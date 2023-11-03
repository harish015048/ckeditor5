using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;

namespace CMCai.Actions
{
    public class Encryption
    {
        static byte[] bytes = ASCIIEncoding.ASCII.GetBytes(ConfigurationManager.AppSettings["PwdKey"].ToString());

        #region EncryptData()
        /// <summary>
        /// Method :: EncryptData
        /// To encrypt Password
        /// </summary>
        /// <param name="strKey">Key</param>
        /// <param name="strData">Password</param>
        /// <returns></returns>

        public string EncryptData(String Input)
        {
            if (String.IsNullOrEmpty(Input))
            {
                throw new ArgumentNullException
                       ("The string which needs to be encrypted can not be null.");
            }
            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream();
            CryptoStream cryptoStream = new CryptoStream(memoryStream,
                cryptoProvider.CreateEncryptor(bytes, bytes), CryptoStreamMode.Write);
            StreamWriter writer = new StreamWriter(cryptoStream);
            writer.Write(Input);
            writer.Flush();
            cryptoStream.FlushFinalBlock();
            writer.Flush();
            return Convert.ToBase64String(memoryStream.GetBuffer(), 0, (int)memoryStream.Length);
        }
        #endregion


        #region DecryptData()
        public string DecryptData(String strData)
        {
            if (String.IsNullOrEmpty(strData))
            {
                throw new ArgumentNullException
                   ("The string which needs to be decrypted can not be null.");
            }
            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream
                    (Convert.FromBase64String(strData));
            CryptoStream cryptoStream = new CryptoStream(memoryStream,
                cryptoProvider.CreateDecryptor(bytes, bytes), CryptoStreamMode.Read);
            StreamReader reader = new StreamReader(cryptoStream);
            return reader.ReadToEnd();
        }
        #endregion
    }
}