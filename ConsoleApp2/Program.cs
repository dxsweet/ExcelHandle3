using System.Diagnostics;
using System.Text;
using Excel1 = Microsoft.Office.Interop.Excel;

internal class Excelhandle2
{
    static void Main(String[] args)
    {
        foreach (Process clsProcess in Process.GetProcesses())
        {
            if (clsProcess.ProcessName.Equals("EXCEL"))
            {
                clsProcess.Kill();
            }
        }

        Excel1.Application xlApp = new Excel1.Application();
        xlApp.Visible = false;
        xlApp.DisplayAlerts = false;

        DirectoryInfo di = new DirectoryInfo(@"./sd");
        foreach (FileInfo fi in di.GetFiles())
        {
            Excel1.Workbook wb = xlApp.Workbooks.Open(fi.FullName);
            Excel1.Worksheet ws = wb.Worksheets["风向风速"];

            StringBuilder  sb1 = new StringBuilder();
            StringBuilder  sb2 = new StringBuilder();

            for (int i = 6; i <= 15 ; i++)
            {
                for(int j = 2; j <= 48; j += 2)
                {
                    if (ws.Cells[i,j].Value2 == null)
                    {
                        sb1.AppendLine("null");
                    }
                    else
                    {
                        sb1.AppendLine(ws.Cells[i,j].Value2.ToString());
                    }

                    if (ws.Cells[i, j + 1].Value2 == null)
                    {
                        sb2.AppendLine("null");
                    }
                    else
                    {
                        sb2.AppendLine(ws.Cells[i, j + 1].Value2.ToString());
                    }

                }
            }


            for (int i = 18; i <= 27; i++)
            {
                for (int j = 2; j <= 48; j += 2)
                {
                    if (ws.Cells[i, j].Value2 == null)
                    {
                        sb1.AppendLine("null");
                    }
                    else
                    {
                        sb1.AppendLine(ws.Cells[i, j].Value2.ToString());
                    }


                    if (ws.Cells[i, j + 1].Value2 == null)
                    {
                        sb2.AppendLine("null");
                    }
                    else
                    {
                        sb2.AppendLine(ws.Cells[i, j + 1].Value2.ToString());
                    }
                }
            }


            for (int i = 30; i <= 40; i++)
            {
                for (int j = 2; j <= 48; j += 2)
                {
                    if (ws.Cells[i, j].Value2 == null)
                    {
                        sb1.AppendLine("null");
                    }
                    else
                    {
                        sb1.AppendLine(ws.Cells[i, j].Value2.ToString());
                    }


                    if (ws.Cells[i, j + 1].Value2 == null)
                    {
                        sb2.AppendLine("null");
                    }
                    else
                    {
                        sb2.AppendLine(ws.Cells[i, j + 1].Value2.ToString());
                    }
                }
            }

            wb.Close();
            xlApp.Workbooks.Close();
           

            String text1 = sb1.ToString().Trim();
            String text2 = sb2.ToString().Trim();
            sb1.Clear();
            sb2.Clear();
            String[] fins = fi.Name.Split('.');
            String fout1 = fins[0] + "fx.txt";
            String fout2 = fins[0] + "fs.txt";

            System.IO.File.WriteAllText(@"./fx/" + fout1, text1, Encoding.UTF8);
            System.IO.File.WriteAllText(@"./fs/" + fout2, text2, Encoding.UTF8);


            Console.WriteLine(fi.Name + "已经处理完毕");
        }


        xlApp.Quit();
        Console.WriteLine("全部处理完毕");


        foreach (Process clsProcess in Process.GetProcesses())
        {
            if (clsProcess.ProcessName.Equals("EXCEL"))
            {
                clsProcess.Kill();
            }
        }

        Console.ReadLine();
    }
}