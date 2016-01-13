using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using Microsoft.Lync.Model.Conversation.AudioVideo;
using Microsoft.Lync.Model.Conversation.Sharing;
using Microsoft.Lync.Model.Extensibility;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;

namespace LyncVideoSync
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Automation lync;
        IAsyncResult handle;
        private LyncClient client;
        private Conversation meetNowConvo;
        
        static AsyncCallback GetNullAsyncCallback(string Message)
        {
            return new AsyncCallback((o) =>
            {
                Debug.Print("NCB : {0}", Message);
                Debug.Print("NCB : {0}", o.GetType().FullName);
            });
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Start_Click(object sender, RoutedEventArgs e)
        {            
            handle = lync.BeginMeetNow(new AsyncCallback((result) =>
            {
                Debug.Print("Meeting created");
            }), null);


        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            lync = LyncClient.GetAutomation();
            client = LyncClient.GetClient();
            ConversationManager mgr = client.ConversationManager;
            mgr.ConversationAdded += (s, evt) =>
            {
                meetNowConvo = evt.Conversation;
                meetNowConvo.PropertyChanged += (s3, e3) =>
                {
                    Debug.Print("> Conversation.PropertyChanged : Property:{0}, Value:{1}", e3.Property, e3.Value);                    
                };
                ConversationWindow win = lync.GetConversationWindow(meetNowConvo);

                Modality m = meetNowConvo.Modalities[ModalityTypes.AudioVideo];
                            
                AVModality avm = (AVModality)m;

                avm.ModalityStateChanged += (s2, e2) =>
                {
                    Debug.Print("> AVModality.ModalityStateChanged : {0}", e2.NewState);
                    if (e2.NewState == ModalityState.Connected)
                    {
                        Debug.Print("** AV Connected**");
                        ApplicationSharingModality m1 = (ApplicationSharingModality)meetNowConvo.Modalities[ModalityTypes.ApplicationSharing];
                        m1.ModalityStateChanged += (s5, e5) =>
                        {
                            Debug.Print("> ApplicationSharingModality.ModalityChanged : {0}", e5.NewState);

                            if (e5.NewState == ModalityState.Connected)
                            {
                                Debug.Print("** Sharing Ok **");
                                //TODO: start record
                                meetNowConvo.BeginSetProperty(ConversationProperty.NumberOfParticipantsRecording, 1, GetNullAsyncCallback("Conversation.BeginSetProperty"), null);
                            }
                        };
                        
                        m1.BeginShareDesktop(GetNullAsyncCallback("ApplicationSharingModality.BeginShareDesktop"), null);
                        
                    }
                };

                avm.BeginConnect(GetNullAsyncCallback("AVModality.BeginConnect"), null);

            };

            //mgr.ConversationAdded += (s, e0) =>
            //{
            //    Conversation conversation = e0.Conversation;
            //    conversation.PropertyChanged += (s1, e1) =>
            //    {
            //        Debug.Print("> Conversation.PropertyChanged");
            //        _debug(e1);
            //    };

            //    WireModalityChangeEvents(conversation);                
            //};
        }

        private void Stop_Click(object sender, RoutedEventArgs e)
        {
            lync.EndMeetNow(handle);
            meetNowConvo.End();
        }

        public static Dictionary<string, object> _dict(object atype)
        {
            if (atype == null) return new Dictionary<string, object>();
            Type t = atype.GetType();
            System.Reflection.PropertyInfo[] props = t.GetProperties();
            Dictionary<string, object> dict = new Dictionary<string, object>();
            foreach (System.Reflection.PropertyInfo prp in props)
            {
                object value = prp.GetValue(atype, new object[] { });
                dict.Add(prp.Name, value);
            }
            return dict;
        }

        public static void printDict(Dictionary<string, object> dict)
        {
            foreach (var item in dict)
            {
                Debug.Print("\"{0}\": {1}", item.Key, item.Value);
            }
        }

        static void _type(object o)
        {
            Debug.Print(o.GetType().FullName);
        }

        static void _debug(object o)
        {
            _type(o);
            var d = _dict(o);
            printDict(d);
        }

        void WireModalityChangeEvents(Conversation conversation)
        {
            Func<string, EventHandler<ModalityStateChangedEventArgs>> makeHandler = new Func<string, EventHandler<ModalityStateChangedEventArgs>>((msg) =>
            {
                return new EventHandler<ModalityStateChangedEventArgs>((sender, evt) =>
                {
                    Debug.Print("> {0} : {1}", msg, evt.NewState);                    
                });
            });

            foreach (var item in conversation.Modalities)
            {
                ModalityTypes modalityType = item.Key;
                Modality m = conversation.Modalities[modalityType];
                if(m != null)
                {
                    string msg = string.Format("({0}) [{1}]", modalityType, m.GetType().FullName);
                    m.ModalityStateChanged += makeHandler(msg);
                }
                else
                {
                    Debug.Print("** Skipped : {0}", modalityType);
                }
            }
        }
    }
}
