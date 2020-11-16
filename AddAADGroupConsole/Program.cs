using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;

namespace AddAADGroupConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            string orgUrl = "";
            Console.WriteLine("Organization Url");
            orgUrl = Console.ReadLine();
            string fileName = "";
            Console.WriteLine("File Path:");
            fileName = Console.ReadLine();

            AddAADGroupToProjectsFromFile(orgUrl, fileName);

            Console.WriteLine("Press enter to exit...");
            Console.ReadLine();
        }

        public static void AddAADGroupToProjectsFromFile(string OrgUrl, string fileName)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileName);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                string teamProject = "";
                string aadGroupName = "";
                string customTpGroupName = "";
                for (int j = 1; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        if (j == 1)
                            teamProject = xlRange.Cells[i, j].Value2.ToString();
                        else if (j == 2)
                            customTpGroupName = xlRange.Cells[i, j].Value2.ToString();
                        else if (j == 3)
                            aadGroupName = xlRange.Cells[i, j].Value2.ToString();
                    }
                }

                AddAADGroupToTPCustomGroup(teamProject, customTpGroupName, aadGroupName, OrgUrl);
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        public static bool AddAADGroupToTPCustomGroup(string teamProject, string tpCustomGroupName, string aadGroupName, string organizationUrl)
        {
            VssCredentials creds = new VssClientCredentials();
            creds.Storage = new VssClientCredentialStorage();

            var tpc = new TfsTeamProjectCollection(new Uri(organizationUrl), creds);
            tpc.Connect(Microsoft.TeamFoundation.Framework.Common.ConnectOptions.IncludeServices);

            IIdentityManagementService ims = tpc.GetService<IIdentityManagementService>();

            string tpCustomGroupNameFull = "[" + teamProject + "]" + "\\" + tpCustomGroupName;
            string aadGroupNameFull = "[TEAM FOUNDATION]" + "\\" + aadGroupName;  //for AAD Groups

            try
            {
                var tfsGroupIdentity = ims.ReadIdentity(IdentitySearchFactor.AccountName,
                                                    tpCustomGroupNameFull,
                                                    MembershipQuery.None,
                                                    ReadIdentityOptions.IncludeReadFromSource);

                var aadGroupIdentity = ims.ReadIdentity(IdentitySearchFactor.AccountName,
                                                        aadGroupNameFull,
                                                        MembershipQuery.None,
                                                        ReadIdentityOptions.IncludeReadFromSource);

                ims.AddMemberToApplicationGroup(tfsGroupIdentity.Descriptor, aadGroupIdentity.Descriptor);

                Console.WriteLine("Group added: " + aadGroupName + " to " + tpCustomGroupName);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Group cannot be added: " + aadGroupName + ", " + tpCustomGroupName);
                return false;
            }
        }
    }
}
