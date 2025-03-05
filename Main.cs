using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using PaperlessQc.Class;

namespace PaperlessQc
{
    public partial class Main : Form
    {
        private FormDirectory directory = new FormDirectory();
        private bool isDebug = Properties.Settings.Default.IsDebug;
        private string dorPass = Properties.Settings.Default.DorPass;
        ModelValidator validator = new ModelValidator();

        public Main()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtModel.Text.Trim()))
                {
                    MessageBox.Show("Please scan ID tag or type product model.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.ActiveControl = txtModel;
                    return;
                }

                string model;
                string qty;
                string lotNo;
                string[] arrSplit;

                arrSplit = txtModel.Text.Trim().Split(null);

                if (arrSplit.Length == 3) // If scanned, extract model, qty, lotNo
                {
                    model = arrSplit[0].ToString();
                    qty = arrSplit[1].ToString();
                    lotNo = arrSplit[2].ToString();
                    txtModel.Text = arrSplit[0].ToString().Trim();
                }
                else
                {
                    model = txtModel.Text.Trim();
                    qty = string.Empty;
                    lotNo = string.Empty;
                }

                if (!validator.IsModelValid(model))
                {
                    MessageBox.Show("Invalid model.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.ActiveControl = txtModel;
                    return;
                }

                string targetDir = directory.DirLiveForm(dtpDate.Value);

                if (!Directory.Exists(targetDir))
                {
                    MessageBox.Show("DOR folder does not exist.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Get all files in the target directory
                string[] allFiles = Directory.GetFiles(targetDir, "*.*");

                // Find all matching inspection performance and packaging checksheet files
                var inspectionFilesA = allFiles
                    .Where(f => Path.GetFileNameWithoutExtension(f).Contains($"{model} Inspection Performance A"))
                    .ToList();

                var inspectionFilesB = allFiles
                .Where(f => Path.GetFileNameWithoutExtension(f).Contains($"{model} Inspection Performance B"))
                .ToList();

                var packagingFiles = allFiles
                    .Where(f => Path.GetFileNameWithoutExtension(f).Contains($"{model} Packaging Checksheet"))
                    .ToList();

                // Track missing files
                List<string> missingFiles = new List<string>();

                if (inspectionFilesA.Count == 0)
                    missingFiles.Add($"{model} Inspection Performance A");

                if (inspectionFilesB.Count == 0)
                    missingFiles.Add($"{model} Inspection Performance B");

                if (packagingFiles.Count == 0)
                    missingFiles.Add($"{model} Packaging Checksheet");

                if (missingFiles.Count > 0)
                {
                    MessageBox.Show($"The following file(s) were not found:\n\n{string.Join("\n", missingFiles)}",
                                    "Missing Files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                // If no files are found, return
                if (inspectionFilesA.Count == 0 && inspectionFilesB.Count == 0 && packagingFiles.Count == 0)
                {
                    return;
                }

                // Disable the main form
                this.Enabled = false;

                Task.Run(() =>
                {
                // Open each existing file
                List<Process> processes = new List<Process>();

                    foreach (var file in inspectionFilesA.Concat(inspectionFilesB).Concat(packagingFiles)) // Open ALL matching files
                {

                    // Copy file to temp directory to avoid locking the original
                    string tempFile = Path.Combine(Path.GetTempPath(), Path.GetFileName(file));
                    File.Copy(file, tempFile, true);

                    using (ExcelPackage package = new ExcelPackage(new FileInfo(tempFile)))
                    {
                        ExcelWorkbook workbook = package.Workbook;

                        // Get last visible sheet that is NOT "Data"
                        ExcelWorksheet lastSheet = workbook.Worksheets.Reverse()
                            .FirstOrDefault(sheet => sheet.Name != "Data");

                        if (lastSheet != null)
                        {
                            lastSheet.Select(); // Set the last non-"Data" sheet as active

                            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(file);

                            if (fileNameWithoutExt.Contains("Inspection Performance A"))
                            {
                                ExcelRange range = lastSheet.Cells["D6:P6"];
                                var lastCell = range.Where(cell => cell.Value != null).LastOrDefault();
                                if (lastCell != null)
                                    lastSheet.View.ActiveCell = lastCell.Address;
                            }
                            else if (fileNameWithoutExt.Contains("Inspection Performance B"))
                                {
                                    ExcelRange range = lastSheet.Cells["D6:P6"];
                                    var lastCell = range.Where(cell => cell.Value != null).LastOrDefault();
                                    if (lastCell != null)
                                        lastSheet.View.ActiveCell = lastCell.Address;
                                }
                            else if (fileNameWithoutExt.Contains("Packaging Checksheet"))
                            {
                                ExcelRange range = lastSheet.Cells["E7:Q7"];
                                var lastCell = range.Where(cell => cell.Value != null).LastOrDefault();
                                if (lastCell != null)
                                    lastSheet.View.ActiveCell = lastCell.Address;
                            }
                        }

                        package.Save();
                    }

                    // Open file and add process to the list
                    Process prc = new Process();
                    prc.StartInfo.FileName = file;
                    prc.StartInfo.UseShellExecute = true;
                    prc.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
                    prc.Start();
                    processes.Add(prc);
                }

                // Wait for all opened files to be closed
                foreach (var prc in processes)
                {
                    prc.WaitForExit();
                }

                    // Re-enable the form after all files are closed (UI update on main thread)
                    this.Invoke((MethodInvoker)delegate {
                        this.Enabled = true;
                    });
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Enabled = true; // Ensure form is re-enabled in case of an error
            }
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            dtpDate.Value = DateTime.Now;
            SetShift();
            this.Text += " " + Application.ProductVersion.ToString();

            string hostname = WindowsIdentity.GetCurrent().Name.Split('\\')[0];
            if (hostname.ToString().Length >= 15) //check if one of the NBPVSPACESERVER
            {
                if (hostname.ToString().Substring(0, 15) == "NBPVSPACESERVER") //if yes
                {
                    lblLine.Text = Environment.UserName.ToString().Trim();
                }
                else //if no
                {
                    lblLine.Text = Environment.MachineName.ToString().Trim();
                }
            }
            else
            {
                lblLine.Text = Environment.MachineName.ToString().Trim();
            }

            txtModel.Text = "7L0053-7025A";
            this.ActiveControl = txtModel;
        }

        private void frmMain_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                e.Handled = true;
                switch (e.KeyCode)
                {
                    case Keys.D1:
                    case Keys.NumPad1:
                        btnSearch.PerformClick();
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtModel_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (!validator.IsModelValid(txtModel.Text.Trim()))
                    {
                        MessageBox.Show("Invalid model.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.ActiveControl = txtModel;
                        btnCreate.Enabled = false;
                        btnSearch.Enabled = false;
                        return;
                    }
                    else
                    {
                        btnCreate.Enabled = true;
                        btnSearch.Enabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetShift()
        {
            string shift = string.Empty;

            try
            {
                if (rdDs.Checked == true)
                    shift = "DS";
                else
                    shift = "NS";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return shift;
        }

        private void SetShift()
        {
            try
            {
                if (DateTime.Now.Hour >= 7 & DateTime.Now.Hour <= 19)
                    rdDs.Checked = true;
                else
                    rdNs.Checked = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtModel.Text.Trim()))
                {
                    MessageBox.Show("Please scan ID tag or type product model.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.ActiveControl = txtModel;
                    return;
                }

                string model;
                string qty;
                string lotNo;
                string[] arrSplit;

                arrSplit = txtModel.Text.Trim().Split(null);

                if (arrSplit.Length == 3) // Prod id tag was scanned
                {
                    model = arrSplit[0].ToString();
                    qty = arrSplit[1].ToString();
                    lotNo = arrSplit[2].ToString();
                    txtModel.Text = arrSplit[0].ToString().Trim();
                }
                else
                {
                    model = txtModel.Text.Trim();
                    qty = string.Empty;
                    lotNo = string.Empty;
                }

                if (!validator.IsModelValid(model))
                {
                    MessageBox.Show("Invalid model.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.ActiveControl = txtModel;
                    return;
                }

                if (!Directory.Exists(directory.DirTemplates()))
                {
                    MessageBox.Show("Templates folder does not exists.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!Directory.Exists(directory.DirLiveForm(dtpDate.Value)))
                {
                    MessageBox.Show("DOR folder does not exists.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string[] templates = Directory.GetFiles(directory.DirTemplates(), "*.*");

                // Get files that match either "Inspection Performance" or "Packaging Checklist"
                var inspectionFilesA = templates
                .Where(f => Path.GetFileNameWithoutExtension(f).Contains($"{model} Inspection Performance A"))
                .ToList();

                var inspectionFilesB = templates
                .Where(f => Path.GetFileNameWithoutExtension(f).Contains($"{model} Inspection Performance B"))
                .ToList();

                var packagingFiles = templates
                    .Where(f => Path.GetFileNameWithoutExtension(f).Contains($"{model} Packaging Checksheet"))
                    .ToList();

                if (inspectionFilesA.Count == 0)
                {
                    MessageBox.Show(model + " Inspection Performance A not found.");
                    return;
                }

                if (inspectionFilesB.Count == 0)
                {
                    MessageBox.Show(model + " Inspection Performance B not found.");
                    return;
                }

                if (packagingFiles.Count == 0)
                {
                    MessageBox.Show(model + " Packaging Checksheet not found.");
                    return;
                }

                // Combine both lists
                var allTemplates = inspectionFilesA.Concat(inspectionFilesB).Concat(packagingFiles).ToList();

                // Find all files that already exist in the target directory
                var existingFiles = allTemplates.Where(file => File.Exists(Path.Combine(directory.DirLiveForm(dtpDate.Value), Path.GetFileName(file)))).ToList();

                if (existingFiles.Count > 0)
                {
                    // Build a message listing all existing files
                    string fileList = string.Join("\n", existingFiles.Select(f => Path.GetFileName(f)));

                    MessageBox.Show($"The following file(s) already exist:\n\n{fileList}\n\nForm creation was canceled.","", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                this.Enabled = false;

                List<Process> processes = new List<Process>();

                foreach (var tempFile in allTemplates)
                {
                    string destinationPath = Path.Combine(directory.DirLiveForm(dtpDate.Value), Path.GetFileName(tempFile));

                    File.Copy(tempFile, destinationPath, false);

                    FileInfo fileInfo = new FileInfo(destinationPath);
                    using (ExcelPackage package = new ExcelPackage(fileInfo))
                    {
                        ExcelWorkbook workbook = package.Workbook;

                        // Since this project targets .NET 4.5, index starts with 1
                        ExcelWorksheet sheet = package.Workbook.Worksheets[1]; // Get first sheet

                        // Unprotect the sheet using the password
                        string sheetPassword = dorPass;  // Sheet password
                        sheet.Protection.SetPassword(sheetPassword); // Unprotect the sheet
                        sheet.Protection.IsProtected = false;

                        // Move horizontal scroll bar one column left
                        sheet.View.FreezePanes(1, 2);
                        sheet.View.UnFreezePanes();

                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(tempFile);

                        if (fileNameWithoutExt.Contains("Inspection Performance A"))
                        {
                            sheet.View.ZoomScale = 100;
                            sheet.Cells[6, 4].Value = dtpDate.Value.ToString("MMddyy");
                            sheet.View.ActiveCell = "D7";

                        } else if (fileNameWithoutExt.Contains("Inspection Performance B"))
                        {
                            sheet.View.ZoomScale = 100;
                            sheet.Cells[6, 4].Value = dtpDate.Value.ToString("MMddyy");
                            sheet.View.ActiveCell = "D7";
                        }
                        else if (fileNameWithoutExt.Contains("Packaging Checksheet"))
                        {
                            sheet.View.ZoomScale = 100;
                            sheet.Cells[7, 5].Value = dtpDate.Value.ToString("MMddyy");
                            sheet.View.ActiveCell = "E8";
                        }

                        // Reapply sheet protection with the same password
                        sheet.Protection.SetPassword(sheetPassword);
                        sheet.Protection.IsProtected = true;

                        // Save the file while preserving macros
                        package.Save();
                    }

                    Process prc = new Process();
                    prc.StartInfo.FileName = destinationPath;
                    prc.StartInfo.UseShellExecute = true; // Allows OS to decide the correct app
                    prc.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;

                    prc.Start();
                    processes.Add(prc); // Store process
                }

                // Wait for all processes to exit
                Task.Run(() =>
                {
                    foreach (var process in processes)
                    {
                        process.WaitForExit();
                    }

                    // Enable the form after all files are closed (UI update on main thread)
                    this.Invoke((MethodInvoker)delegate {
                        this.Enabled = true;
                    });
                });

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
