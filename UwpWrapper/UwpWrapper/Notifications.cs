using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Windows.Data.Xml.Dom;
using Windows.UI.Notifications;

namespace UwpWrapper
{
    [Guid("D28468EF-3769-4D5C-9505-92577937CEB9")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [ComVisible(true)]
    public interface INotifications
    {
        void LaunchNotification(string title);

        void LaunchNotification2(string imageTitle, string imagePath);

        void LaunchNotificationCustom(string imageTitle, string imagePath);
    }

    [Guid("F419465C-76B7-46FD-A1F5-DD3601E3FF02")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    [ProgId("UwpWrapper.Notifications")]
    public class Notifications : INotifications
    {
        void INotifications.LaunchNotification(string title)
        {
            ToastTemplateType toastTemplate = ToastTemplateType.ToastText02;
            XmlDocument toastXml = ToastNotificationManager.GetTemplateContent(toastTemplate);

            XmlNodeList toastTextElements = toastXml.GetElementsByTagName("text");
            toastTextElements[0].AppendChild(toastXml.CreateTextNode(title));
            toastTextElements[1].AppendChild(toastXml.CreateTextNode(DateTime.Now.ToString()));

            ToastNotification toast = new ToastNotification(toastXml);
            ToastNotificationManager.CreateToastNotifier().Show(toast);
        }

        void INotifications.LaunchNotification2(string title, string imagePath)
        {
            ToastTemplateType toastTemplate = ToastTemplateType.ToastImageAndText02;
            XmlDocument toastXml = ToastNotificationManager.GetTemplateContent(toastTemplate);

            XmlNodeList toastImageAttributes = toastXml.GetElementsByTagName("image");
            ((XmlElement)toastImageAttributes[0]).SetAttribute("src", imagePath);
            ((XmlElement)toastImageAttributes[0]).SetAttribute("alt", title);

            ToastNotification toast = new ToastNotification(toastXml);
            ToastNotificationManager.CreateToastNotifier().Show(toast);
        }

        void INotifications.LaunchNotificationCustom(string title, string imagePath)
        {
            string toastTemplate = "<toast launch='" + imagePath + "'><visual><binding template='ToastGeneric'><text>" + title + "</text><image src='" + imagePath + "' /></binding></visual></toast>";
            XmlDocument toastXml = new XmlDocument();
            toastXml.LoadXml(toastTemplate);
            ToastNotification toast = new ToastNotification(toastXml);
            ToastNotificationManager.CreateToastNotifier().Show(toast);
        }
    }
}
