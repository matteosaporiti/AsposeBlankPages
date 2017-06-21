using System;
using System.IO;


namespace AsposeBlankPages
{
    public static class AsposeLicenseManager
    {
        public static void RegWordsLibrary(LicenceType License, string licensePath, bool ThrowExIfExpired)
        {
            try
            {
                License.ApplyLicense(licensePath, ThrowExIfExpired, lic =>
                {
                    global::Aspose.Words.License TypedLicense = new global::Aspose.Words.License();
                    if (TypedLicense != null)
                    {

                        TypedLicense.SetLicense(lic);
                    }
                    else
                    {
                        throw new Exception("Typed Licence Aspose.Words is null");
                    }
                });
            }
            catch (System.Exception ex)
            {
                throw new System.Exception("Generic Exception on Create and Set Aspose.Words.License", ex);
            }
        }

        public delegate void CodeToApplyLicense(MemoryStream lic);

        private static void ApplyLicense(this LicenceType License, string licensePath, bool ThrowExIfExpired, CodeToApplyLicense ApplyLicenseCode)
        {
            try
            {
                using (Lic lic = new Lic(License, licensePath))
                {
                    if (lic != null)
                    {
                        if (lic.CanApply)
                        {
                            if (ApplyLicenseCode != null)
                            {
                                ApplyLicenseCode(lic.LicenseStream);
                            }
                            else
                            {
                                throw new Exception("Delegate ApplyLicenseCode of type CodeToApplyLicense is null");
                            }
                        }
                    }
                }
            }
            catch (InvalidOperationException IoEx)
            {
                if (IoEx.Message.ToLowerInvariant().Contains("expired"))
                {
                    if (ThrowExIfExpired)
                    {
                        throw new Exception("The Aspose license has expired.", IoEx);
                    }
                }
                else
                {
                    throw new Exception("InvalidOperationException on applying Aspose license!", IoEx);
                }
            }
            catch (System.Exception ex)
            {
                throw new Exception("Generic Exception on applying Aspose license!", ex);
            }
        }

        internal class Lic : IDisposable
        {
            public bool CanApply = true;
            public MemoryStream LicenseStream;

            public Lic(LicenceType License, string licensePath)
            {
                try
                {
                    switch (License)
                    {
                        case LicenceType.Production:
                            LicenseStream = new MemoryStream(File.ReadAllBytes(licensePath));
                            if (CheckLicenceStream(LicenseStream, "Aspose_Total"))
                            {
                                CanApply = true;
                            }
                            break;
                    }
                }
                catch (System.Exception ex)
                {
                    CanApply = false;
                    throw new System.Exception("Loading Aspose license error!", ex);
                }

            }
            public static bool CheckLicenceStream(MemoryStream Mem, string ResourceName)
            {
                bool result = true;
                if (Mem == null)
                {
                    result = false;
                    throw new Exception(string.Format("License Stream \"{0}\" is null", ResourceName));
                }
                else if (Mem.Length <= 0)
                {
                    result = false;
                    throw new Exception(string.Format("License Stream \"{0}\" Length is {1}", ResourceName, Mem.Length));
                }
                return result;
            }
            public void Dispose()
            {
                if (LicenseStream != null)
                {
                    LicenseStream.Close();
                }
                if (LicenseStream != null)
                {
                    LicenseStream.Dispose();
                }
            }
        }

        public enum LicenceType
        {
            NoLicense = 0,
            Production = 1,
            TemporaryUnlimited = 2,
            Expired = 3
        }
    }
}
