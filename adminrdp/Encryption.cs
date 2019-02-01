/*
Created by: John Walker
Last modified: 1/31/2018
Description: Encrypts and decrypts passwords for saved credentials. There are plans to implement DPAPI at a later date
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.Collections;
using System.IO;
using System.Management;
using System.Windows.Forms;

namespace AdminRDP
{
    class Encryption
    {

        //generate passphrase based on off of wmi
        private static string Grab_PassPhrase()
        {
            try
            {
                string results = null;
                string results2 = null;
                string passphrase = null;
                string query = "SELECT Manufacturer, Name FROM Win32_ComputerSystem";
                string query2 = "SELECT SerialNumber FROM Win32_BIOS";
                ManagementObjectSearcher moSearch = new ManagementObjectSearcher(query);
                ManagementObjectSearcher moSearch2 = new ManagementObjectSearcher(query2);
                ManagementObjectCollection moCollection = moSearch.Get();
                ManagementObjectCollection moCollection2 = moSearch2.Get();

                foreach (ManagementObject mo in moCollection)
                {
                    results = mo["Name"].ToString() + mo["Manufacturer"].ToString();
                    results = results.Replace(" ", String.Empty);
                }

                foreach (ManagementObject mo2 in moCollection2)
                {
                    results2 = mo2["SerialNumber"].ToString();
                }
                passphrase = results2 + results;
                return passphrase;
            }catch(Exception e)
            {
                return MessageBox.Show(e.Message).ToString();
            }
        }
        // This size of the IV (in bytes) must = (keysize / 8).  Default keysize is 256, so the IV must be
        // 32 bytes long.  Using a 16 character string here gives us 32 bytes when converted to a byte array.

        // This constant is used to determine the keysize of the encryption algorithm
        private static int keysize = 256;
        private static string initVector = "2v4d0v76d9sw0fk3";
        private static string passPhrase = Grab_PassPhrase();

        //Encrypt
        public static string EncryptString (string plainText)
        {
            try
            {
                byte[] initVectorBytes = Encoding.UTF8.GetBytes(initVector);
                byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);
                PasswordDeriveBytes password = new PasswordDeriveBytes(passPhrase, null);
                byte[] keyBytes = password.GetBytes(keysize / 8);
                RijndaelManaged symmetricKey = new RijndaelManaged();
                symmetricKey.Mode = CipherMode.CBC;
                ICryptoTransform encryptor = symmetricKey.CreateEncryptor(keyBytes, initVectorBytes);
                MemoryStream memoryStream = new MemoryStream();
                CryptoStream cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write);
                cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
                cryptoStream.FlushFinalBlock();
                byte[] cipherTextBytes = memoryStream.ToArray();
                memoryStream.Close();
                cryptoStream.Close();
                return Convert.ToBase64String(cipherTextBytes);
            }catch(Exception e2)
            {
                return MessageBox.Show(e2.Message).ToString();
            }

        }
        //Decrypt
        public static string DecryptString(string cipherText)
        {
            try
            {
                byte[] initVectorBytes = Encoding.UTF8.GetBytes(initVector);
                byte[] cipherTextBytes = Convert.FromBase64String(cipherText);
                PasswordDeriveBytes password = new PasswordDeriveBytes(passPhrase, null);
                byte[] keyBytes = password.GetBytes(keysize / 8);
                RijndaelManaged symmetricKey = new RijndaelManaged();
                symmetricKey.Mode = CipherMode.CBC;
                ICryptoTransform decryptor = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes);
                MemoryStream memoryStream = new MemoryStream(cipherTextBytes);
                CryptoStream cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
                byte[] plainTextBytes = new byte[cipherTextBytes.Length];
                int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
                memoryStream.Close();
                cryptoStream.Close();
                return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount);
            }catch(Exception e3)
            {
                return MessageBox.Show(e3.Message).ToString();
            }
        }


    }
}
