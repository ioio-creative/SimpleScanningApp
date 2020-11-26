using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using WIA;

namespace SimpleScanningApp
{
    public partial class Form1 : Form
    {
        private const string ScannerName = "EPSON Perfection V37/V370";
        private const string OutputPath = @"E:/ScanImg.png";
        private const int ScannerDpiHorizontal = 300; //1200;
        private const int ScannerDpiVertical = 300; //1200;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // https://youtu.be/PwnCjZSnE_E
            try
            {
                foreach (DeviceInfo deviceInfo in GetScannerDeviceInfos())
                {
                    lstListOfScanners.Items.Add(GetNameFromDeviceInfo(deviceInfo));
                }
            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnScan_Click(object sender, EventArgs e)
        {
            // https://youtu.be/PwnCjZSnE_E
            // https://www.codesenior.com/sources/docs/tutorials/How-to-get-Images-From-Scanner-in-C-with-Windows-Image-AcqusitionWIA-Library-.pdf
            try
            {            
                DeviceInfo availableScanner = GetScannerDeviceByName(ScannerName);                

                if (availableScanner == null)
                {
                    MessageBox.Show("Scanner does not exist");
                    return;
                }

                var device = availableScanner.Connect();
                //device.Properties.get_Item("3088").set_Value(5); // Double scanning or dublex scanning
                //device.Properties.get_Item("3088").set_Value(1); // Single scanning or simplex scanning

                var scannerItem = device.Items[1];
                scannerItem.Properties.get_Item("6147").set_Value(ScannerDpiHorizontal); // horizontal dpi
                scannerItem.Properties.get_Item("6148").set_Value(ScannerDpiVertical); // vertical dpi

                var imgFile = scannerItem.Transfer(FormatID.wiaFormatJPEG) as ImageFile;
                var path = OutputPath;

                if (File.Exists(path))
                {
                    File.Delete(path);
                }


                imgFile.SaveFile(path);

                pictureBox1.ImageLocation = path;
            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private static IEnumerable<DeviceInfo> GetScannerDeviceInfos()
        {
            return new DeviceManager().DeviceInfos.OfType<DeviceInfo>().Where(x => x.Type == WiaDeviceType.ScannerDeviceType);
        }

        private static DeviceInfo GetScannerDeviceByName(string name)
        {
            string cleanedName = name.ToUpper();
            return GetScannerDeviceInfos().FirstOrDefault(x => GetNameFromDeviceInfo(x).ToUpper().Contains(cleanedName));
        }

        private static string GetNameFromDeviceInfo(DeviceInfo deviceInfo)
        {
            return deviceInfo.Properties["Name"].get_Value() as string;
        }
    }
}
