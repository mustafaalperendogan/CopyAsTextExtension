using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Specialized;
using System.ComponentModel.Design;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace CopyAsTextExtension
{
    internal sealed class CopyFileAsTextCommand
    {
        // VSCT'deki ID'ler:
        public const int CommandId_ItemNode = 0x0100; // CopyFileAsText_ItemNode
        public const int CommandId_CodeWin = 0x0101; // CopyFileAsText_CodeWin

        public static readonly Guid CommandSet = new Guid("a8b5c3d1-2f4e-4a7b-8c9d-1e3f5a7b9c2d");

        private readonly AsyncPackage package;

        private CopyFileAsTextCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            this.package = package ?? throw new ArgumentNullException(nameof(package));
            if (commandService == null) throw new ArgumentNullException(nameof(commandService));

            // 1) Solution Explorer bağlamı
            var cmdIdItem = new CommandID(CommandSet, CommandId_ItemNode);
            var cmdItem = new OleMenuCommand(Execute, cmdIdItem);
            cmdItem.BeforeQueryStatus += OnBeforeQueryStatus;
            commandService.AddCommand(cmdItem);

            // 2) Editör bağlamı
            var cmdIdCode = new CommandID(CommandSet, CommandId_CodeWin);
            var cmdCode = new OleMenuCommand(Execute, cmdIdCode);
            cmdCode.BeforeQueryStatus += OnBeforeQueryStatus;
            commandService.AddCommand(cmdCode);
        }

        public static CopyFileAsTextCommand Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new CopyFileAsTextCommand(package, commandService);
        }

        private void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (sender is OleMenuCommand cmd)
            {
                var path = GetSelectedFilePath();
                cmd.Enabled = !string.IsNullOrEmpty(path) && File.Exists(path);
                cmd.Visible = true; // istersen görünürlük de togglenabilir
            }
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                string selectedFilePath = GetSelectedFilePath();

                if (string.IsNullOrEmpty(selectedFilePath) || !File.Exists(selectedFilePath))
                {
                    ShowMessage("Lütfen bir dosya seçin veya editörde açık bir dosya bulundurun.", "Uyarı");
                    return;
                }

                string fileContent = File.ReadAllText(selectedFilePath);

                string fileName = Path.GetFileNameWithoutExtension(selectedFilePath) + ".txt";
                string tempPath = Path.Combine(Path.GetTempPath(), "VSExtension_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tempPath);
                string tempFilePath = Path.Combine(tempPath, fileName);

                File.WriteAllText(tempFilePath, fileContent);

                if (!CopyFileToClipboard(tempFilePath))
                {
                    ShowMessage("Dosya kopyalama başarısız oldu. Lütfen tekrar deneyin.", "Hata");
                    try { Directory.Delete(tempPath, true); } catch { }
                    return;
                }

                ShowMessage($"'{fileName}' hafızaya kopyalandı. Yapıştırabilirsiniz.", "Başarılı");
            }
            catch (Exception ex)
            {
                ShowMessage($"Hata oluştu: {ex.Message}", "Hata");
            }
        }

        private bool CopyFileToClipboard(string filePath)
        {
            bool success = false;
            Exception lastException = null;

            var thread = new Thread(() =>
            {
                try
                {
                    for (int i = 0; i < 3; i++)
                    {
                        try
                        {
                            var files = new StringCollection { filePath };
                            Clipboard.Clear();
                            Thread.Sleep(10);
                            Clipboard.SetFileDropList(files);
                            success = true;
                            break;
                        }
                        catch (Exception ex)
                        {
                            lastException = ex;
                            Thread.Sleep(50);
                        }
                    }
                }
                catch (Exception ex)
                {
                    lastException = ex;
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(5000);

            if (!success && lastException != null)
                System.Diagnostics.Debug.WriteLine($"Clipboard hatası: {lastException.Message}");

            return success;
        }

        private string GetSelectedFilePath()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = Package.GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE.DTE;

            // 1) Solution Explorer'da seçili dosya
            if (dte?.SelectedItems?.Count > 0)
            {
                foreach (EnvDTE.SelectedItem item in dte.SelectedItems)
                {
                    if (item?.ProjectItem?.Properties != null)
                    {
                        try
                        {
                            var fullPathObj = item.ProjectItem.Properties.Item("FullPath")?.Value;
                            var path = fullPathObj?.ToString();
                            if (!string.IsNullOrEmpty(path))
                                return path;
                        }
                        catch { /* bazı item tiplerinde FullPath olmayabilir */ }
                    }
                }
            }

            // 2) Aktif editör
            if (dte?.ActiveDocument != null)
                return dte.ActiveDocument.FullName;

            return null;
        }

        private void ShowMessage(string message, string title)
        {
            VsShellUtilities.ShowMessageBox(
                package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
