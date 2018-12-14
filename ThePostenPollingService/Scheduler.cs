using Digipost.Signature.Api.Client.Core;
using Digipost.Signature.Api.Client.Core.Exceptions;
using Digipost.Signature.Api.Client.Portal;
using Digipost.Signature.Api.Client.Portal.Enums;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Timers;
using Therefore.API;

namespace ThePostenPollingService
{
    public partial class Scheduler : ServiceBase
    {
        //The Schdeuler which makes sure the void in timer1_Tick runs ever defined seconds
        private Timer timer1 = null;

        public Scheduler()
        {
            InitializeComponent();
        }

        //An event that occurs every time the service starts (holds the ticker)
        protected override void OnStart(string[] args)
        {
            timer1 = new Timer();
            this.timer1.Interval = 120000; // Evry minute
            this.timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.timer1_Tick);
            timer1.Enabled = true;
            Library.WriteErrorLog("The Posten Polling Service started");

        }


        private void timer1_Tick(object sender, ElapsedEventArgs e)
        {

            X509Certificate2 tekniskAvsenderSertifikat;
            string scertificateThumbprint ="";
            string sorganizationNumber = "";
            string sEnvironment = "";
            string sDebugMode = "";
            int iDebugMode = 0;
            string iterations = "";

            //regSubKeycatNos contains a list if all category Numbers, JobID-field name for the category and status for the category for all workflows.
            List<string> regSubKeyCatNos = new List<string>();

            //Load the values from the registry
            RegistryKey key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
            key = key.OpenSubKey("SOFTWARE").OpenSubKey("Canon");

            if (key != null)
            {

                sorganizationNumber = key.GetValue("OrganizationNo").ToString();
                scertificateThumbprint = key.GetValue("CertThumbprint").ToString();
                sEnvironment = key.GetValue("Environment").ToString();
                sDebugMode = key.GetValue("DebugMode").ToString();
                iterations = key.GetValue("NumberOfPolls").ToString();
                
                if(sDebugMode != "")
                    iDebugMode = Convert.ToInt32(sDebugMode);
                

                foreach (string wfName in key.GetSubKeyNames())
                {
                    if (wfName == "OrganizationNo" || wfName == "CertThumbprint" || wfName == "PollingInterval" || wfName == "NumberOfPolls" || wfName == "Environment" || wfName == "DebugMode")
                        continue;

                    RegistryKey tempkey;
                    tempkey = key.OpenSubKey(wfName);
                    regSubKeyCatNos.Add(tempkey.GetValue("CategoryNo").ToString() + "|" + tempkey.GetValue("JobIDFieldName").ToString() + "|" + tempkey.GetValue("StatusFieldName").ToString());

                    if (iDebugMode > 0)
                        Library.WriteErrorLog("Loaded workflow: " + tempkey.GetValue("CategoryNo").ToString() + "|" + tempkey.GetValue("JobIDFieldName").ToString() + "|" + tempkey.GetValue("StatusFieldName").ToString());

                }
            }

            else
                Library.WriteErrorLog("No workflow exists/has been configured correctly: No information stored in the registry about the certification or organization number");


            //POSTEN: initiate a portalClient
            PortalClient portalClient = null;
            
            Digipost.Signature.Api.Client.Core.Environment env = Digipost.Signature.Api.Client.Core.Environment.DifiTest;
            if (sEnvironment == "Production")
            {
                env = Digipost.Signature.Api.Client.Core.Environment.Production;
                
                //DEBUG
                if (iDebugMode > 0)
                    Library.WriteErrorLog("!!!!!!!!!!!!Working against the prdouction environment !!!!!!");
            }

            tekniskAvsenderSertifikat = getCertificat(scertificateThumbprint);

            //POSTEN: Create Client Configuration
            ClientConfiguration clientconfiguration = new ClientConfiguration(
               env,
                tekniskAvsenderSertifikat,
                new Sender(sorganizationNumber));

            //DEBUG
            if(iDebugMode > 0){ 
            Library.WriteErrorLog("Client configured: " + clientconfiguration.ToString());
            Library.WriteErrorLog("Current Environment: " + clientconfiguration.Environment.ToString());
                }

            portalClient = new PortalClient(clientconfiguration);

            //POSTEN: Do the initial poll and get a status
            JobStatusChanged jobStatusChanged = portalClient.GetStatusChange().Result;

            if (jobStatusChanged.Status == JobStatus.NoChanges)
            {
                //Queue is empty. Additional polling will result in blocking for a defined period.
                //DEBUG
                if (iDebugMode > 0)
                    Library.WriteErrorLog("First Poll: No Change.." + " ");
            }
            else
            {
                //DEBUG
                if (iDebugMode > 0)
                    Library.WriteErrorLog("Staus changed - polling jobs");

                for (int i = 0; i < Convert.ToInt32(iterations); i++)
                {
                    if (i != 0)
                        jobStatusChanged = portalClient.GetStatusChange().Result;

                    try
                    {
                        JobStatus signatureJobStatus = jobStatusChanged.Status;
                        string docNo = jobStatusChanged.JobReference;
                        string tmpFileName = "";

                        if (signatureJobStatus == JobStatus.CompletedSuccessfully)
                        {
                            //DEBUG
                            if (iDebugMode > 1)
                                Library.WriteErrorLog("Polling job: " + jobStatusChanged.JobId);

                            Stream pades = portalClient.GetPades(jobStatusChanged.PadesReference).Result;
                            tmpFileName = System.IO.Path.GetTempPath() + jobStatusChanged.JobId + "_"+ Guid.NewGuid().ToString()  + ".pdf";
                           
                            

                            FileStream filestream = File.Create(tmpFileName, (int)pades.Length);
                            byte[] bytesInStream = new byte[pades.Length];
                            pades.Read(bytesInStream, 0, bytesInStream.Length);
                            filestream.Write(bytesInStream, 0, bytesInStream.Length);

                            filestream.Close();
                            filestream.Dispose();

                            //DEBUG
                            if (iDebugMode > 1)
                                Library.WriteErrorLog("Polling job: " + jobStatusChanged.JobId + ": Pades retrieved");
                        }


                        TheServer server = new TheServer();
                        server.Connect(TheClientType.CustomService);
                        //DEBUG
                        if (iDebugMode > 0)
                            Library.WriteErrorLog("Connected to The Server");


                    
                        TheIndexData theInData = new TheIndexData();
                        TheDocument theDoc = new TheDocument();
                        bool fileSaved = false;

                        //DEBUG
                        if (iDebugMode > 0)
                            Library.WriteErrorLog("Document Number from Reference: " + docNo);

                        theInData.DocNo = Convert.ToInt32(docNo);
                        theInData.Load(server);
                        string sPostenStatus = "";
                        string statusFieldName = "";
                        statusFieldName = getStatusFieldName(theInData.CtgryNo.ToString(), regSubKeyCatNos);
                        
                        if (iDebugMode > 0)
                            Library.WriteErrorLog("Found status field name: " + statusFieldName);
                        if (iDebugMode > 0)
                            Library.WriteErrorLog("CatNo found: " + theInData.CtgryNo.ToString());
                        if (signatureJobStatus != JobStatus.CompletedSuccessfully)
                                {
                                    if (signatureJobStatus == JobStatus.Failed)
                                    {
                                        sPostenStatus = "Signature was either rejected or failed";

                                    }
                                    else if (signatureJobStatus == JobStatus.InProgress)
                                    {
                                        List<Signature> signStatus = new List<Signature>();
                                        signStatus = jobStatusChanged.Signatures;
                                        foreach (Signature sign in signStatus)
                                        {
                                            string tempMail = sign.Identifier.ToContactInformation().Email.Address;
                                    
                                        }

                                        sPostenStatus = "Signature is in Progress";
                                    }
                             
                            if (iDebugMode > 0)
                                Library.WriteErrorLog("Updating status for document : " + docNo);
                            if (!string.IsNullOrEmpty(statusFieldName))
                            {
                                theInData.SetValueByColName(statusFieldName, sPostenStatus);
                            }
                            theInData.SaveChanges(server);

                            //DEBUG
                            if (iDebugMode > 0)
                                Library.WriteErrorLog("Status updated for document: " + docNo);

                        }
                        else {
                            if (iDebugMode > 0)
                               Library.WriteErrorLog("Begin archiving the doc : " + docNo.ToString());

                               TheDocument newDoc = new TheDocument();
                               newDoc.Create("");
                               newDoc.AddStream(tmpFileName, "", 0);
                               newDoc.IndexData = theInData;
                            if (!string.IsNullOrEmpty(statusFieldName))
                            {
                                newDoc.IndexData.SetValueByColName(statusFieldName, "Signed");
                            }

                               string testUser;
                               int testCurrVersion;
                               newDoc.CheckOut(server, 0, out testUser, out testCurrVersion);
                               newDoc.Archive(server, Convert.ToInt32(docNo), "Added  by The Polling Service");
                               newDoc.Dispose();

                               if (iDebugMode > 0)
                                      Library.WriteErrorLog("Document is archived : " + docNo);

                                }
                               

                      if(iDebugMode > 0)
                            Library.WriteErrorLog("Confirming status for JOb ID (Posten): " + jobStatusChanged.JobId);

                        //Confirming that we have recieved the JobStatusChange
                        string confStatus = portalClient.Confirm(jobStatusChanged.ConfirmationReference).Status.ToString();
                        fileSaved = true;

                       if (iDebugMode > 0)
                            Library.WriteErrorLog("Confirming status for JOb ID (Posten): " + jobStatusChanged.JobId + ":  - Confirmed: " + confStatus);

                       if (fileSaved)
                       {

                            File.Delete(tmpFileName);

                            if (iDebugMode > 0)
                                Library.WriteErrorLog("Deleting Temp-file: " + tmpFileName);
                       }

                    }catch(TooEagerPollingException eagerPollingException)
                    {
                        Library.WriteErrorLog("No more jobs to poll: " + eagerPollingException.ToString());
                        break;
                    }
                }
            }
            }

        //System.IO.Stream xades = portalClient.GetXadejobStatusChanged.Signatures.ElementAt(0).XadesReference).Result;
        //Get PAdES:


        //To clean the Thumbprint for the certificate
        public string CleanThumbprint(string mmcThumbprint)
        {
            //replace spaces, non word chars and convert to uppercase
            return Regex.Replace(mmcThumbprint, @"\s|\W", "").ToUpper();
        }

        //Function for finding the locale certificate given a thumbprint
        private X509Certificate2 getCertificat(string thumbprint)
        {
            //Get the certificate from the locale store
            X509Store storeMy = new X509Store(StoreName.My, StoreLocation.LocalMachine);
            X509Certificate2 tekniskAvsenderSertifikat;
            try
            {
                storeMy.Open(OpenFlags.ReadOnly);
                var certs = storeMy.Certificates;

                tekniskAvsenderSertifikat = certs.Find(
                X509FindType.FindByThumbprint, CleanThumbprint(thumbprint), false)[0];
            }
            finally
            {
                storeMy.Close();
            }
            return tekniskAvsenderSertifikat;
        }
        // return the fieldName for the category
        private string getStatusFieldName(string inCatNo, List<string> regSubKeyCatNos)
        {
            string sStatusFieldName = "";
            foreach (string sCatNo in regSubKeyCatNos)
            {
                if (sCatNo.Split(new char[] { '|' })[0].Trim() == inCatNo)
                {
                    sStatusFieldName = sCatNo.Split(new char[] { '|' })[2];
                }
            }
            return sStatusFieldName;
        }

        protected override void OnStop()
        {
            timer1.Enabled = false;
            Library.WriteErrorLog("The Posten Polling Service stopped");
        }
    }
}
