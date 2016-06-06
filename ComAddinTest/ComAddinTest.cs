using System;
using System.Runtime.InteropServices;
using VBIDE = NetOffice.VBIDEApi;
using NetOffice.Tools;
using System.Windows.Forms;

namespace ComAddinTest
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid("5DAEE010-E9DF-4F1A-83C1-1FD088596108")]
    public interface IComAddinTest
    {
        void TestMethod();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("0DD5D44E-E7A0-4991-B8DB-7518C3B99B28")]
    //[ProgId("ComAddinTest.ComAddinTest")]
    public class ComAddinTest : IComAddinTest, IDTExtensibility2
    {
        private VBIDE.Window m_Window;
        private IDEForm m_IdeForm;

        public void TestMethod()
        {

        }

        public void OnAddInsUpdate(ref Array custom)
        {
            
        }

        public void OnBeginShutdown(ref Array custom)
        {
            
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            m_IdeForm = new IDEForm();
            using (var app = (VBIDE.VBE)
                NetOffice.Core.Default.CreateKnownObjectFromComProxy(null, Application, typeof(VBIDE.VBE)))

            using (var addIn = (VBIDE.AddIn)
                    NetOffice.Core.Default.CreateKnownObjectFromComProxy(null, AddInInst, typeof(VBIDE.AddIn)))
            {
                m_Window = app.Windows.CreateToolWindow(addIn, "ComAddinTest.IDEForm", "COM Add-in Test", "EA40F1AF-76EE-49A6-A707-6C08BEB3F46B",
                    m_IdeForm);

                m_Window.Visible = true;
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            //m_Window.Dispose();
            //m_IdeForm.Dispose();
        }

        public void OnStartupComplete(ref Array custom)
        {
            
        }
    }
}
