using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace DemoAddonsDisplayFileExcel
{
    class Common
    {


        public static string ToXMLString<T>(object data)
        {
            using (var stringwriter = new System.IO.StringWriter())
            {
                var serializer = new XmlSerializer(typeof(T));
                serializer.Serialize(stringwriter, data);
                return stringwriter.ToString();
            }
        }


        public static string folderPath()
        {
            string path = ""; NativeWindow nativewindow = new NativeWindow(); Thread t = new Thread(() =>
            {
                nativewindow.AssignHandle(System.Diagnostics.Process.GetProcessesByName("SAP Business One")[0].MainWindowHandle);
                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog(); DialogResult dr = folderBrowserDialog.ShowDialog(nativewindow);
                if (dr == DialogResult.OK)
                {
                    path = folderBrowserDialog.SelectedPath;
                }
            });          // Kick off a new thread
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
            return path;
        }
        public static string openFileDialog(string title, string filter)
        {
            string filename = "";

            // chua thay su khac biet cua nativewindow
            NativeWindow nativewindow = new NativeWindow();

            //chua hieu lam
            Thread t = new Thread(() =>
            {
                nativewindow.AssignHandle(System.Diagnostics.Process.GetProcessesByName("SAP Business One")[0].MainWindowHandle);
                OpenFileDialog openFileDialog = new OpenFileDialog(); openFileDialog.Title = title;
                openFileDialog.Filter = filter;
                DialogResult dr = openFileDialog.ShowDialog(nativewindow);
                if (dr == DialogResult.OK)
                {
                    filename = openFileDialog.FileName;
                }
            });          // Kick off a new thread
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join(); return filename;
        }
        public static string OpenFolderDialog()
        {
            string folderPath = "";
            OpenFileDialog folderBrowser = new OpenFileDialog();
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                folderPath = Path.GetDirectoryName(folderBrowser.FileName);
            }
            return folderPath;
        }
        public static string saveFileDialog(string title, string filter)
        {
            string filename = "";
            NativeWindow nativewindow = new NativeWindow(); Thread t = new Thread(() =>
            {
                nativewindow.AssignHandle(System.Diagnostics.Process.GetProcessesByName("SAP Business One")[0].MainWindowHandle);
                SaveFileDialog saveFileDialog = new SaveFileDialog(); saveFileDialog.Title = title;
                saveFileDialog.Filter = filter; DialogResult dr = saveFileDialog.ShowDialog(nativewindow);
                if (dr == DialogResult.OK)
                {
                    filename = saveFileDialog.FileName;
                }
            });          // Kick off a new thread
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join(); return filename;
        }
    }
}
