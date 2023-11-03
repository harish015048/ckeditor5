using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Text.RegularExpressions;
using CMCai.Models;
using Aspose.Words;
using Aspose.Words.Layout;

namespace CMCai.Actions
{
    public class FolderActions
    {
        // check folder contains sub folders or not check
        public void NoSubFolderCheck(RegOpsQC rObj)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string[] folders = Directory.GetDirectories(rObj.FolderPath);
                string folder = new DirectoryInfo(rObj.FolderPath).Name;
                rObj.Folder_Name = folder;
                if (folders.Length > 0)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Sub folders exists in '" + folder + "'";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        // check Maximum Folder Lenght check
        public void MaximumFolderLenghtCheck(RegOpsQC rObj)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string folder = new DirectoryInfo(rObj.FolderPath).Name;
                rObj.Folder_Name = folder;
                if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                {
                    if (folder.Length > Convert.ToInt64(rObj.Check_Parameter.ToString()))
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Folder length exceeded given characters limit of : " + rObj.Check_Parameter + "";
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Folder length not defined ";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        // Folder Naming Convention Check
        public void FolderNamingConventionCheck(RegOpsQC rObj, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            string DescritivePrefix = string.Empty;
            try
            {
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                DirectoryInfo folder = new DirectoryInfo(rObj.FolderPath);
                Regex regexData = new Regex(@"([^a-zA-Z0-9])", RegexOptions.IgnoreCase);
                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[k].JID = rObj.JID;
                        chLst[k].Job_ID = rObj.Job_ID;
                        chLst[k].Folder_Name = rObj.Folder_Name;
                        chLst[k].Created_ID = rObj.Created_ID;
                        chLst[k].FolderPath = rObj.FolderPath;
                        chLst[k].File_Name = null;

                        if (chLst.Count == 1 && chLst[k].Check_Name != "Accepted Special Characters")
                        {
                            if (!regexData.IsMatch(folder.Name))
                            {
                                rObj.QC_Result = "Passed";
                            }
                            else
                            {
                                rObj.QC_Result = "Failed";
                                rObj.Comments = "Folder name contains special character(s)";
                            }
                        }
                        if (chLst[k].Check_Name == "Accepted Special Characters")
                        {
                            try
                            {
                                Regex regexData1 = null;

                                if (chLst[k].Check_Parameter != null)
                                    chLst[k].Check_Parameter = chLst[k].Check_Parameter.Replace("[", "").Replace("\"", "").Replace("]", "").Replace("\\", "");

                                if (chLst[k].Check_Parameter.ToString() == "Underscore")
                                {
                                    chLst[k].Check_Parameter = "_";
                                }
                                if (chLst[k].Check_Parameter.ToString() == "Underscore,Hypen")
                                {
                                    chLst[k].Check_Parameter = "_-";

                                }
                                if (chLst[k].Check_Parameter.ToString() == "Hypen")
                                {
                                    chLst[k].Check_Parameter = "-";
                                }
                                DescritivePrefix = chLst[k].Check_Parameter;
                                regexData1 = new Regex(@"([^a-zA-Z0-9" + DescritivePrefix + "])", RegexOptions.IgnoreCase);

                                if (!regexData1.IsMatch(folder.Name))
                                {
                                    if (rObj.Check_Type == 0)
                                    {
                                        rObj.QC_Result = "Passed";
                                    }
                                    else
                                    {
                                        rObj.QC_Result = "Failed";
                                        rObj.Comments = "Folder name contains special character(s)";
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Failed";
                                    rObj.Comments = "Folder name contains special character(s)";
                                }
                                rObj.CHECK_END_TIME = DateTime.Now;
                            }

                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        // Folder Naming Convention Fix
        public void FolderNamingConventionFix(RegOpsQC rObj, List<RegOpsQC> chLst)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixFlag = false;
            bool AcceptFlag = false;
            string DescritivePrefix = string.Empty;
            List<string> splCharInFolder = new List<string>();
            List<string> splCharInFile = new List<string>();
            string NewFolderName = string.Empty;
            string NewFileName = string.Empty;
            try
            {
                // to get sub checks list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                DirectoryInfo folder = new DirectoryInfo(rObj.FolderPath);
                FileInfo[] files = folder.GetFiles();
                Regex regexData = new Regex(@"([^a-zA-Z0-9])", RegexOptions.IgnoreCase);
                foreach (char c in folder.Name)
                {
                    if (regexData.IsMatch(c.ToString()))
                    {
                        if (!splCharInFolder.Contains(c.ToString()))
                        {
                            splCharInFolder.Add(c.ToString());
                        }
                    }
                }
                string acceptsplchar = string.Empty;
                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[k].JID = rObj.JID;
                        chLst[k].Job_ID = rObj.Job_ID;
                        chLst[k].Folder_Name = rObj.Folder_Name;
                        chLst[k].Created_ID = rObj.Created_ID;
                        chLst[k].FolderPath = rObj.FolderPath;

                        if (chLst[k].Check_Name == "Accepted Special Characters")
                        {
                            acceptsplchar = chLst[k].Check_Parameter;
                        }
                        if (chLst[k].Check_Name == "Replace Special Characters")
                        {
                            try
                            {
                                if (chLst[k].Check_Parameter == "Remove")
                                {
                                    chLst[k].Check_Parameter = "";
                                }
                                else if (chLst[k].Check_Parameter == "Underscore")
                                {
                                    chLst[k].Check_Parameter = "_";
                                }
                                else if (chLst[k].Check_Parameter == "Hypen")
                                {
                                    chLst[k].Check_Parameter = "-";
                                }
                                NewFolderName = folder.Name;
                                foreach (string c in splCharInFolder)
                                {
                                    FixFlag = true;
                                    if (!acceptsplchar.Contains(c))
                                    {
                                        NewFolderName = NewFolderName.Replace(c, chLst[k].Check_Parameter);
                                    }
                                }
                                // rename new folder name 
                                string[] folderNewPath2 = rObj.FolderPath.Split(new string[] { folder.Name }, StringSplitOptions.None);
                                string oldFolder = folderNewPath2[0] + folder.Name;
                                string newFolder = folderNewPath2[0] + NewFolderName;
                                if (oldFolder != newFolder)
                                {
                                    Directory.Move(oldFolder, newFolder);
                                }
                                rObj.Folder_Name = NewFolderName;
                                rObj.FolderPath = newFolder;
                            }
                            catch (Exception ex)
                            {
                                rObj.QC_Result = "Error";
                                rObj.Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                    }
                }
                for (int l = 0; l < chLst.Count; l++)
                {
                    chLst[l].Folder_Name = NewFolderName;
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                }         
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        // Folder Name Prefix Check
        public void FolderNamePrefixCheck(RegOpsQC rObj, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty; 
            bool DescriptivePrefix = false;
            bool ConsecutiveNum = false;
            bool desPrefix = false;
            bool numPrefix = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            string consecutiveNumPrefix = string.Empty;
            string DescritivePrefix = string.Empty;
            string finaltest = string.Empty;
            try
            {
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                string folder = new DirectoryInfo(rObj.FolderPath).Name;
                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[k].JID = rObj.JID;
                        chLst[k].Job_ID = rObj.Job_ID;
                        chLst[k].Folder_Name = folder;
                        chLst[k].Created_ID = rObj.Created_ID;
                        chLst[k].FolderPath = rObj.FolderPath;

                        if (chLst[k].Check_Name == "Descriptive Prefix")
                        {                           
                                DescritivePrefix = chLst[k].Check_Parameter;
                                desPrefix = true;                            
                        }
                        else if (chLst[k].Check_Name == "Consecutive Number Prefix")
                        {
                            string newFolder = string.Empty;
                            string num1 = string.Empty;
                            //if (rObj.folderNum == 1)
                            //{
                            //    consecutiveNumPrefix = chLst[k].Check_Parameter;
                            //}
                            //else
                            //{
                            //    foreach (char c in chLst[k].Check_Parameter)
                            //    {
                            //        if (c.ToString() != "0")
                            //        {
                            //            newFolder = rObj.folderNum.ToString();
                            //            num1 = chLst[k].Check_Parameter.Replace(c.ToString(), newFolder);
                            //        }
                            //    }
                            //    consecutiveNumPrefix = num1;
                            //}
                            if (rObj.folderNum == 1)
                            {
                                num1 = chLst[k].Check_Parameter.ToString();
                            }
                            else
                            {                                
                                num1 = rObj.folderNum.ToString("D" + chLst[k].Check_Parameter.ToString().Length);
                            }                            
                            consecutiveNumPrefix = num1;
                            numPrefix = true;                            
                        }
                    }
                }
                if (numPrefix == true && desPrefix == true)
                {
                    if (folder.StartsWith(DescritivePrefix))
                    {
                        DescriptivePrefix = true;
                    }
                    if (folder.StartsWith(consecutiveNumPrefix))
                    {
                        ConsecutiveNum = true;
                    }
                    if (DescriptivePrefix == false && ConsecutiveNum == false)
                    {
                        finaltest = DescritivePrefix + "_" + consecutiveNumPrefix + "_" + folder;
                    }
                    else if (DescriptivePrefix == false && ConsecutiveNum == true)
                    {
                        finaltest = DescritivePrefix + "_" + folder;
                    }
                    else if (DescriptivePrefix == true && ConsecutiveNum == false)
                    {
                        finaltest = consecutiveNumPrefix + "_" + folder;
                    }                    
                    if (folder.StartsWith(finaltest))
                    {
                        rObj.QC_Result = "Passed";
                    }
                    else
                    {                        
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Folder name does not contains " + finaltest;
                    }
                }
                if (numPrefix == true && desPrefix == false)
                {
                     finaltest = consecutiveNumPrefix + "_" + folder;
                    if (folder.StartsWith(consecutiveNumPrefix))
                    {
                        rObj.QC_Result = "Passed";
                    }
                    else
                    {                        
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Folder name does not contains " + consecutiveNumPrefix;
                    }
                }
                if (desPrefix == true && numPrefix == false)
                {
                     finaltest = DescritivePrefix + "_" + folder;
                    if (folder.StartsWith(DescritivePrefix))
                    {
                        rObj.QC_Result = "Passed";
                    }
                    else
                    {                        
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Folder name does not contains " + DescritivePrefix;

                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        // Folder Name Prefix Fix
        public void FolderNamePrefixFix(RegOpsQC rObj, List<RegOpsQC> chLst)
        {
            string res = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixFlag = false;
            bool DescriptivePrefix = false;
            bool ConsecutiveNum = false;
            bool desPrefix = false;
            bool numPrefix = false;
            string consecutiveNumPrefix = string.Empty;
            string DescritivePrefix = string.Empty;
            string finaltest = string.Empty;

            try
            {
                // to get sub checks list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                string folder = new DirectoryInfo(rObj.FolderPath).Name;
                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[k].JID = rObj.JID;
                        chLst[k].Job_ID = rObj.Job_ID;
                        chLst[k].Folder_Name = folder;
                        chLst[k].File_Name = rObj.File_Name;
                        chLst[k].Created_ID = rObj.Created_ID;
                        chLst[k].FolderPath = rObj.FolderPath;
                        if (chLst[k].Check_Name == "Descriptive Prefix")
                        {
                            try
                            {
                                DescritivePrefix = chLst[k].Check_Parameter;
                                desPrefix = true;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Consecutive Number Prefix")
                        {
                            try
                            {
                                string newFolder = string.Empty;
                                string num1 = string.Empty;
                                //if (rObj.folderNum == 1)
                                //{
                                //    consecutiveNumPrefix = chLst[k].Check_Parameter;
                                //}
                                //else
                                //{
                                //    foreach (char c in chLst[k].Check_Parameter)
                                //    {
                                //        if (c.ToString() != "0")
                                //        {
                                //            newFolder = rObj.folderNum.ToString();
                                //            num1 = chLst[k].Check_Parameter.Replace(c.ToString(), newFolder);
                                //        }
                                //    }
                                //    consecutiveNumPrefix = num1;
                                //}

                                if (rObj.folderNum == 1)
                                {
                                    num1 = chLst[k].Check_Parameter.ToString();
                                }
                                else
                                {                                    
                                    num1 = rObj.folderNum.ToString("D" + chLst[k].Check_Parameter.ToString().Length);
                                }
                                consecutiveNumPrefix = num1;                               
                                numPrefix = true;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }

                    }
                }

                if (numPrefix == true && desPrefix == true)
                {
                    if (folder.StartsWith(DescritivePrefix))
                    {
                        DescriptivePrefix = true;
                    }
                   else if (folder.StartsWith(consecutiveNumPrefix))
                   {
                        ConsecutiveNum = true;
                   }
                    if (DescriptivePrefix == false && ConsecutiveNum == false)
                    {
                        finaltest = DescritivePrefix + "_" + consecutiveNumPrefix + "_" + folder;
                    }
                    else if (DescriptivePrefix == false && ConsecutiveNum == true)
                    {
                        finaltest = DescritivePrefix + "_" + folder;
                    }
                    else if (DescriptivePrefix == true && ConsecutiveNum == false)
                    {
                        finaltest = consecutiveNumPrefix + "_" + folder;
                    }
                    if (folder.Contains(finaltest))
                    {
                        rObj.QC_Result = "Passed";
                    }
                    else
                    {
                        FixFlag = true;
                        string NewFolderName = finaltest;
                        string[] folderNewPath2 = rObj.FolderPath.Split(new string[] { folder }, StringSplitOptions.None);
                        string oldFolder = folderNewPath2[0] + folder;
                        string newFolder = folderNewPath2[0] + NewFolderName;
                        if (oldFolder != newFolder)
                        {
                            Directory.Move(oldFolder, newFolder);
                        }    
                        rObj.Folder_Name = finaltest;
                        rObj.FolderPath = newFolder;
                    }

                }
                if (numPrefix == true && desPrefix == false)
                {
                    finaltest = consecutiveNumPrefix + "_" + folder;
                    if (folder.StartsWith(finaltest))
                    {
                        rObj.QC_Result = "Passed";
                    }
                    else
                    {
                        FixFlag = true;
                        string NewFolderName = finaltest;
                        string[] folderNewPath2 = rObj.FolderPath.Split(new string[] { folder }, StringSplitOptions.None);
                        string oldFolder = folderNewPath2[0] + folder;
                        string newFolder = folderNewPath2[0] + NewFolderName;
                        if (oldFolder != newFolder)
                        {
                            Directory.Move(oldFolder, newFolder);
                        }
                        rObj.Folder_Name = finaltest;
                        rObj.FolderPath = newFolder;

                    }
                }
                if (desPrefix == true && numPrefix == false)
                {
                    finaltest = DescritivePrefix + "_" + folder;
                    if (folder.StartsWith(finaltest))
                    {
                        rObj.QC_Result = "Passed";
                    }
                    else
                    {
                        FixFlag = true;
                        string NewFolderName = finaltest + " _" + folder;
                        string[] folderNewPath2 = rObj.FolderPath.Split(new string[] { folder }, StringSplitOptions.None);
                        string oldFolder = folderNewPath2[0] + folder;
                        string newFolder = folderNewPath2[0] + NewFolderName;
                        if (oldFolder != newFolder)
                        {
                            Directory.Move(oldFolder, newFolder);
                        }
                        rObj.Folder_Name = finaltest;
                        rObj.FolderPath = newFolder;
                    }
                }
                for (int l = 0; l < chLst.Count; l++)
                {
                    chLst[l].Folder_Name = finaltest;
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + " Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        static int GetIntValueFromString(string input)
        {
            var result = 0;
            var intString = Regex.Replace(input, "[^0-9]+", string.Empty);
            Int32.TryParse(intString, out result);
            return result;
        }

        // Folder Sequence Files Naming Check
        public void FolderSequenceNameCheck(RegOpsQC rObj)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;           
            rObj.CHECK_START_TIME = DateTime.Now;
            string FileName = string.Empty;
            try
            {
                DirectoryInfo folder = new DirectoryInfo(rObj.FolderPath);
                FileInfo[] files = folder.GetFiles();
                if (files.Length > 0)
                {
                    List<string> NumberfileNames = new List<string>();
                    List<Int32> NumberfileNumbes = new List<int>();
                    List<Int32> NumberfileNumbes1 = new List<int>();
                    List<string> NonNumberfileNames = new List<string>();
                    List<Int32> NonNumberfileNumbes = new List<int>();
                    int num222 = 0;
                    foreach (var file1 in files)
                    {
                        num222 = GetIntValueFromString(file1.Name);
                        if (num222 > 0)
                        {
                            NumberfileNumbes.Add(num222);
                            NumberfileNumbes1.Add(num222);
                            NumberfileNames.Add(file1.Name);
                        }
                        else
                        {
                            NonNumberfileNumbes.Add(num222);
                            NonNumberfileNames.Add(file1.Name);
                        }

                    }
                    NumberfileNumbes.Sort();

                    int number = 0;

                    if (NumberfileNumbes.Count > 0)
                    {
                        foreach (var file in NumberfileNumbes)
                        {
                            if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                            {
                                if (number == 0)
                                {
                                    number = Convert.ToInt32(rObj.Check_Parameter.ToString());
                                }
                                else
                                {
                                    number++;
                                }
                                string str = number.ToString("D" + rObj.Check_Parameter.Length);
                                int index111 = NumberfileNumbes1.IndexOf(file);
                                string filenum = NumberfileNames[index111];
                                FileName = str + "_" + filenum;
                                if (!filenum.StartsWith(str))
                                {
                                    rObj.QC_Result = "Failed";
                                    rObj.Comments = "Files in the Folder with given number is not set to " + rObj.Check_Parameter + "";
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                            else
                            {
                                rObj.QC_Result = "Failed";
                                rObj.Comments = "Files in the Folder with given number is not defined";
                            }
                        }
                    }
                    if (NonNumberfileNumbes.Count > 0)
                    {
                        foreach (var file in NonNumberfileNames)
                        {
                            if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                            {
                                if (number == 0)
                                {
                                    number = Convert.ToInt32(rObj.Check_Parameter.ToString());
                                }
                                else
                                {
                                    number++;
                                }
                                string str = number.ToString("D" + rObj.Check_Parameter.Length);
                                //int index111 = NonNumberfileNumbes.IndexOf(file);
                                //string filenum = NonNumberfileNames[index111];
                                FileName = str + "_" + file;
                                if (!file.StartsWith(str))
                                {
                                    rObj.QC_Result = "Failed";
                                    rObj.Comments = "Files in the Folder with given number is not set to " + rObj.Check_Parameter + "";
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                            else
                            {
                                rObj.QC_Result = "Failed";
                                rObj.Comments = "Files in the Folder with given number is not defined";
                            }
                        }
                    }

                }
                
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        // Folder Sequence Files Naming Fix
        public void FolderSequenceNameFix(RegOpsQC rObj, List<RegOpsQC> chLst)
        {
            string res = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixFlag = false;
            List<string> filesList = new List<string>();
            string FileName = string.Empty;
            try
            {
                if (chLst.Count > 0)
                {
                    if (rObj.Check_Name == "Sequential Files Naming")
                    {
                        try
                        {
                            DirectoryInfo folder = new DirectoryInfo(rObj.FolderPath);
                            FileInfo[] files = folder.GetFiles();
                            if (files.Length > 0)
                            {
                                List<string> NumberfileNames = new List<string>();
                                List<Int32> NumberfileNumbes = new List<int>();
                                List<Int32> NumberfileNumbes1 = new List<int>();
                                List<string> NonNumberfileNames = new List<string>();
                                List<Int32> NonNumberfileNumbes = new List<int>();
                                int num222 = 0;
                                foreach (var file1 in files)
                                {
                                    num222 = GetIntValueFromString(file1.Name);
                                    if (num222 > 0)
                                    {
                                        NumberfileNumbes.Add(num222);
                                        NumberfileNumbes1.Add(num222);
                                        NumberfileNames.Add(file1.Name);
                                    }
                                    else
                                    {
                                        NonNumberfileNumbes.Add(num222);
                                        NonNumberfileNames.Add(file1.Name);
                                    }

                                }
                                NumberfileNumbes.Sort();

                                rObj.Folder_Name = folder.Name;
                                int number = 0;
                                string file33 = string.Empty;

                                if (NumberfileNumbes.Count > 0)
                                {
                                    foreach (var file in NumberfileNumbes)
                                    {
                                        if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                                        {
                                            int index111 = NumberfileNumbes1.IndexOf(file);
                                            string filenum = NumberfileNames[index111];

                                            if (!filenum.StartsWith(rObj.Check_Parameter.ToString()))
                                            {
                                                FixFlag = true;
                                                if (number == 0)
                                                {
                                                    number = Convert.ToInt32(rObj.Check_Parameter.ToString());
                                                }
                                                else
                                                {
                                                    number++;
                                                }
                                                string str = number.ToString("D" + rObj.Check_Parameter.Length);
                                                FileName = str + "_" + filenum;
                                                string oldpath = rObj.FolderPath + "//" + filenum;
                                                string newpath = rObj.FolderPath + "//" + FileName;
                                                File.Move(oldpath, newpath);
                                            }
                                            else
                                            {
                                                rObj.QC_Result = "Passed";
                                            }
                                        }
                                        filesList.Add(FileName);

                                    }
                                }
                                if (NonNumberfileNumbes.Count > 0)
                                {
                                    foreach (var file in NonNumberfileNames)
                                    {
                                        if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                                        {
                                            if (!file.StartsWith(rObj.Check_Parameter.ToString()))
                                            {
                                                FixFlag = true;
                                                if (number == 0)
                                                {
                                                    number = Convert.ToInt32(rObj.Check_Parameter.ToString());
                                                }
                                                else
                                                {
                                                    number++;
                                                }
                                                string str = number.ToString("D" + rObj.Check_Parameter.Length);
                                                FileName = str + "_" + file;
                                                string oldpath = rObj.FolderPath + "//" + file;
                                                string newpath = rObj.FolderPath + "//" + FileName;
                                                File.Move(oldpath, newpath);
                                            }
                                            else
                                            {
                                                rObj.QC_Result = "Passed";
                                            }
                                        }
                                        filesList.Add(FileName);

                                    }
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                            rObj.QC_Result = "Error";
                            rObj.Comments = "Technical error: " + ex.Message;
                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                        }
                    }
                }
                string[] strFile = filesList.ToArray();
                string filelst = string.Empty;
                for (int j = 0; j < strFile.Count(); j++)
                {
                    filelst += strFile[j] + ", ";
                }
                if (filelst.EndsWith(", "))
                {
                    filelst = filelst.Remove(filelst.Length - 2);
                }
                if (FixFlag == true)
                {

                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + " For The list of Files " + filelst + " Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        // File Naming Convention Fix
        public void FileNamingConvention(RegOpsQC rObj, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixFlag = false;
            bool AcceptFlag = false;
            string consecutiveNumPrefix = string.Empty;
            string DescritivePrefix = string.Empty;            
            List<string> splCharInFile = new List<string>();
            string NewFolderName = string.Empty;
            string NewFileName = string.Empty;
            string fileNameWithoutExtension = string.Empty;
            string extension = string.Empty;
            try
            {
                // to get sub checks list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                Regex regexData = new Regex(@"([^a-zA-Z0-9])", RegexOptions.IgnoreCase);
                string acceptsplchar = string.Empty;
                rObj.JID = rObj.JID;
                fileNameWithoutExtension = Path.GetFileNameWithoutExtension(rObj.File_Name);
                extension = Path.GetExtension(rObj.File_Name);
                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[k].JID = rObj.JID;
                        chLst[k].Job_ID = rObj.Job_ID;
                        chLst[k].Folder_Name = rObj.Folder_Name;
                        chLst[k].Created_ID = rObj.Created_ID;
                        chLst[k].FolderPath = rObj.FolderPath;
                        if (chLst[k].Check_Name == "Accepted Special Characters")
                        {
                            if (chLst[k].Check_Parameter != null)
                                chLst[k].Check_Parameter = chLst[k].Check_Parameter.Replace("[", "").Replace("\"", "").Replace("]", "").Replace("\\", "");

                            if (chLst[k].Check_Parameter.ToString() == "Underscore")
                            {
                                chLst[k].Check_Parameter = "_";
                                chLst[k].CheckAcceptedParamVal = "_";
                                rObj.CheckAcceptedParamVal = "_";
                            }
                            if (chLst[k].Check_Parameter.ToString() == "Underscore,Hypen")
                            {
                                chLst[k].Check_Parameter = "_-";
                                chLst[k].CheckAcceptedParamVal = "_-";
                                rObj.CheckAcceptedParamVal = "_-";

                            }
                            if (chLst[k].Check_Parameter.ToString() == "Hypen")
                            {
                                chLst[k].Check_Parameter = "-";
                                chLst[k].CheckAcceptedParamVal = "-";
                                rObj.CheckAcceptedParamVal = "-";
                            }
                            acceptsplchar = chLst[k].Check_Parameter;
                        }
                        if (chLst[k].Check_Name == "Accepted Special Characters")
                        {
                            try
                            {
                                DescritivePrefix = chLst[k].Check_Parameter;
                                Regex regexData1 = new Regex(@"([^a-zA-Z0-9" + DescritivePrefix + "])", RegexOptions.IgnoreCase);
                                if (!regexData1.IsMatch(fileNameWithoutExtension))
                                {
                                    rObj.QC_Result = "Passed";
                                }
                                else
                                {
                                    AcceptFlag = true;
                                    rObj.QC_Result = "Failed";
                                    rObj.Comments = "File name does not contains '" + chLst[k].Check_Parameter + "'";
                                }
                            }
                            catch (Exception ex)
                            {
                                rObj.QC_Result = "Error";
                                rObj.Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Replace Special Characters")
                        {
                            try
                            {
                                FixFlag = true;
                                if (chLst[k].Check_Parameter == "Remove")
                                {
                                    chLst[k].Check_Parameter = "";
                                }
                                else if (chLst[k].Check_Parameter == "Underscore" || chLst[k].Check_Parameter == "_")
                                {
                                    chLst[k].Check_Parameter = "_";
                                    chLst[k].CheckParamVal = "_";
                                    rObj.CheckParamVal = "_";
                                }
                                else if (chLst[k].Check_Parameter == "Hypen" || chLst[k].Check_Parameter == "-")
                                {
                                    chLst[k].Check_Parameter = "-";
                                    chLst[k].CheckParamVal = "-";
                                    rObj.CheckParamVal = "-";
                                }

                                // get special characters in file                                
                                foreach (char c in fileNameWithoutExtension)
                                {
                                    if (regexData.IsMatch(c.ToString()))
                                    {
                                        if (!splCharInFile.Contains(c.ToString()))
                                        {
                                            splCharInFile.Add(c.ToString());

                                        }
                                    }
                                }
                                NewFileName = fileNameWithoutExtension;
                                foreach (string c in splCharInFile)
                                {
                                    FixFlag = true;
                                    if (!acceptsplchar.Contains(c))
                                    {
                                        NewFileName = NewFileName.Replace(c, chLst[k].Check_Parameter);
                                    }
                                }
                                // rename new file name 
                                string oldpath = rObj.FolderPath + "//" + rObj.File_Name;
                                string newpath = rObj.FolderPath + "//" + NewFileName + extension;
                                if (oldpath != newpath)
                                {
                                    File.Move(oldpath, newpath);
                                }                                
                                rObj.File_Name = NewFileName + extension;
                                chLst[k].File_Name = NewFileName + extension;
                                rObj.Folder_Name = new DirectoryInfo(rObj.FolderPath).Name;

                            }
                            catch (Exception ex)
                            {
                                rObj.QC_Result = "Error";
                                rObj.Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                    }
                }
                for (int l = 0; l < chLst.Count; l++)
                {
                    chLst[l].Folder_Name = new DirectoryInfo(rObj.FolderPath).Name;
                    chLst[l].File_Name = NewFileName + extension;
                }

                if (FixFlag == true && AcceptFlag == true)
                {

                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + " Fixed";
                }
                else if (AcceptFlag == true)
                {
                    rObj.QC_Result = "Failed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
    }
}