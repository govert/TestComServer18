using System.Diagnostics;
using System.Runtime.InteropServices;

namespace TestComServer18
{
    // Step 1: Defines an event sink interface (MessageEvents) to be implemented by the COM sink.
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface MessageEvents
    {
        void NewMessage(string s);
    }

    // Step 2: Defines a custom interface for the MessageHandler class.
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IMessageHandler
    {
        void FireNewMessageEvent(string s);
    }

    // Step 3: Connects the event sink interface to a class by passing the namespace and event sink interface
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]  // Avoid using AutoDual
    [ComSourceInterfaces(typeof(MessageEvents))]
    public class MessageHandler : IMessageHandler
    {
        public delegate void NewMessageDelegate(string s);
        public event NewMessageDelegate NewMessage;
        public MessageHandler() { }

        public void FireNewMessageEvent(string s)
        {
            Debug.Print($"New Message {s}");
            if (NewMessage != null)
            {
                Debug.Print($"Invoke {s}");
                NewMessage.Invoke(s);
            }
        }
    }
}
