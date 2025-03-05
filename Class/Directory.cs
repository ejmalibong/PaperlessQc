using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PaperlessQc.Class
{
    class FormDirectory
    {
        private bool isDebug = Properties.Settings.Default.IsDebug;

        public string DirDefault()
        {
            string dir = string.Empty;

            try
            {
                if (isDebug == true)
                {
                    dir = @"B:\Users BACKUP\NBCP-LT-144\Desktop\Attachment\Quality";
                }
                else
                {
                    dir = @"\\192.168.20.11\Quality\10_QC\71_DOR";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return dir;
        }

        public string DirFormType(int i)
        {
            string dir = string.Empty;

            try
            {
                switch (i)
                {
                    case 1:
                        dir = @"7L0053-7025A Inspection Performance A";
                        break;

                    case 2:
                        dir = @"7L0053-7025A Inspection Performance B";
                        break;

                    case 3:
                        dir = @"7L0053-7025A Packaging Checksheet";
                        break;

                    default:
                        dir = @"7L0053-7025A Inspection Performance A";
                        break;
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return dir;
        }

        public string DirTemplates()
        {
            string dir = string.Empty;
            
            try
            {
                dir = Path.Combine(DirDefault(), @"TEMPLATES");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return dir;
        }

        public string TemplateName(int i)
        {
            string name = string.Empty;

            try
            {
                switch (i)
                {
                    case 1:
                        name = "Daily Operation Report.xlsx";
                        break;

                    default:
                        name = "Daily Operation Report.xlsx";
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return name;
        }

        public string DirLiveForm(DateTime dt)
        {
            string dir = string.Empty;

            try
            {
                dir = Path.Combine(DirDefault(), @"DOR", dt.ToString("yyyy"), MonthNumber(dt.ToString("MMMM")) + " " + dt.ToString("MMMM"));
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return dir;
        }

        public int MonthNumber(string monthName)
        {
            int monthNo = 0;

            try
            {
                switch (monthName)
                {
                    case "January":
                        monthNo = 1;
                        break;

                    case "February":
                        monthNo = 2;
                        break;

                    case "March":
                        monthNo = 3;
                        break;

                    case "April":
                        monthNo = 4;
                        break;

                    case "May":
                        monthNo = 5;
                        break;

                    case "June":
                        monthNo = 6;
                        break;

                    case "July":
                        monthNo = 7;
                        break;

                    case "August":
                        monthNo = 8;
                        break;

                    case "September":
                        monthNo = 9;
                        break;

                    case "October":
                        monthNo = 10;
                        break;

                    case "November":
                        monthNo = 11;
                        break;

                    case "December":
                        monthNo = 12;
                        break;

                    default:
                        monthNo = 1;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return monthNo;
        }
    }
}
