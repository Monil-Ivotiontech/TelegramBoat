using System.Configuration;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Harsh
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
        // Constants for mouse events
        private const uint MOUSEEVENTF_LEFTDOWN = 0x02;
        private const uint MOUSEEVENTF_LEFTUP = 0x04;
        int x1 = 0; // X position
        int y1 = 0; // Y position
        int x2 = 0; // X position
        int y2 = 0; // Y position
        int x3 = 0; // X position
        int y3 = 0; // Y position
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            int Interval = Convert.ToInt32(ConfigurationManager.AppSettings["Interval"]);
            timerset.Interval = Interval;
            timerset.Enabled = true;
            x1 = Convert.ToInt16(ConfigurationManager.AppSettings["X1"]);
            y1 = Convert.ToInt16(ConfigurationManager.AppSettings["Y1"]);
            x2 = Convert.ToInt16(ConfigurationManager.AppSettings["X2"]);
            y2 = Convert.ToInt16(ConfigurationManager.AppSettings["Y2"]);
            x3 = Convert.ToInt16(ConfigurationManager.AppSettings["X3"]);
            y3 = Convert.ToInt16(ConfigurationManager.AppSettings["Y3"]);
            label1.Text = "Data will download at every " + (Interval / 1000) + " seconds";
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                Cursor.Position = new System.Drawing.Point(x1, y1);
                // Simulate mouse click
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, (uint)x1, (uint)y1, 0, 0);

                Thread.Sleep(15000);

                Cursor.Position = new System.Drawing.Point(x2, y2);
                // Simulate mouse click
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, (uint)x2, (uint)y2, 0, 0);

                Thread.Sleep(15000);
                
                Cursor.Position = new System.Drawing.Point(x3, y3);
                // Simulate mouse click
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, (uint)x3, (uint)y3, 0, 0);

                Thread.Sleep(15000);

                string sourceFilePath = ConfigurationManager.AppSettings["SourceFilePath"];
                // Define the destination file path (the new location for the file)
                string destinationFilePath = ConfigurationManager.AppSettings["DestinationFilePath"];

                label2.Text = "Current source folder " + sourceFilePath;
                label3.Text = "Current destination folder " + sourceFilePath;

                string[] files = Directory.GetFiles(sourceFilePath);

                // Loop through each file and move it to the destination folder
                foreach (string filePath in files)
                {
                    string ext = Path.GetExtension(filePath);
                    if(ext.ToLower().Contains("xls") || ext.ToLower().Contains("xlsx"))
                    {
                        // Get the file name
                        string fileName = Path.GetFileNameWithoutExtension(filePath);

                        // Combine the destination folder with the file name
                        Guid guid = Guid.NewGuid();
                        string newFileName = fileName + "_" + guid.ToString() + ".xlsx";
                        string destinationPath = Path.Combine(destinationFilePath, newFileName);

                        // Move the file
                        File.Move(filePath, destinationPath);

                    }
                }
            }
            catch (Exception ex)
            {

            }

        }
    }
}
