using Microsoft.Win32.Interop;
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace MsiTest
{
    internal class MsiInstaller
    {
        #region msi.dllのプロパティ&メソッド群
        [DllImport("msi.dll", CharSet = CharSet.Unicode, PreserveSig = true, SetLastError = true, ExactSpelling = true)]
        public static extern UInt32 MsiOpenPackageExW(string szPackagePath, UInt32 dwOptions, out IntPtr hProduct);

        [DllImport("msi.dll", SetLastError = true)]
        public static extern int MsiGetProductProperty(IntPtr hProduct, string szProperty, StringBuilder lpValueBuf, ref uint pcchValueBuf);

        [DllImport("msi.dll", ExactSpelling = true)]
        public static extern uint MsiCloseHandle(IntPtr hAny);

        #endregion

        enum MSIOPENPACKAGEFLAGS : uint
        {
            /// <summary>
            /// ignore the machine state when creating the engine
            /// </summary>
            MSIOPENPACKAGEFLAGS_IGNOREMACHINESTATE = 1
        }

        #region Open時のモード

        /// <summary>
        /// 以下の定義を実装するクラス
        /// #define MSIDBOPEN_READONLY     (LPCTSTR)0  // database open read-only, no persistent changes
        /// #define MSIDBOPEN_TRANSACT     (LPCTSTR)1  // database read/write in transaction mode
        /// #define MSIDBOPEN_DIRECT       (LPCTSTR)2  // database direct read/write without transaction
        /// #define MSIDBOPEN_CREATE       (LPCTSTR)3  // create new database, transact mode read/write
        /// #define MSIDBOPEN_CREATEDIRECT (LPCTSTR)4  // create new database, direct mode read/write
        /// #define MSIDBOPEN_PATCHFILE    32/sizeof(*MSIDBOPEN_READONLY) // add flag to indicate patch file
        /// </summary>

        static public class MsiDbOpen
        {
            /// <summary>
            /// 読み取りモード
            /// </summary>
            public static readonly IntPtr READONLY = new IntPtr(0);

            /// <summary>
            /// トランザクションモード
            /// </summary>
            public static readonly IntPtr TRANSACT = new IntPtr(1);

            /// <summary>
            /// トランザクションなしで直接読み書きモード
            /// </summary>
            public static readonly IntPtr DIRECT = new IntPtr(2);

            /// <summary>
            /// 新規作成モード
            /// </summary>
            public static readonly IntPtr CREATE = new IntPtr(3);

            /// <summary>
            /// トランザクションなしで新規作成
            /// </summary>
            public static readonly IntPtr CREATEDIRECT = new IntPtr(4);
        }

        #endregion

        /// <summary>
        /// メジャーバージョン
        /// </summary>
        public int Major { get; protected set; }
        /// <summary>
        /// マイナーバージョン
        /// </summary>
        public int Minor { get; protected set; }
        /// <summary>
        /// ビルドバージョン
        /// </summary>
        public int Build { get; protected set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        private MsiInstaller() { }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="msiPath">MSIファイルのパス</param>
        public MsiInstaller(string msiPath)
        {
            IntPtr hProduct = new IntPtr();
            hProduct = IntPtr.Zero;

            try
            {
                uint winerror = MsiOpenPackageExW(msiPath, 0, out hProduct);
                if (ResultWin32.ERROR_SUCCESS != winerror)
                {
                    throw new MsiInstallerException("MsiOpenPackageExW : " + winerror.ToString());
                }


                StringBuilder sb = new StringBuilder(1024);
                uint capacity = (uint)sb.Capacity;
                winerror = (uint)MsiGetProductProperty(hProduct, @"ProductVersion", sb, ref capacity);
                if (ResultWin32.ERROR_SUCCESS != winerror)
                {
                    throw new MsiInstallerException("MsiGetProductProperty : " + winerror.ToString());
                }

                string productVersion = sb.ToString();

                Match match = Regex.Match(sb.ToString(), @"(?<maj>\d+)\.(?<min>\d+)\.(?<build>\d+)");
                this.Major = int.Parse(match.Groups["maj"].Value);
                this.Minor = int.Parse(match.Groups["min"].Value);
                this.Build = int.Parse(match.Groups["build"].Value);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (hProduct != IntPtr.Zero)
                {
                    MsiCloseHandle(hProduct);
                }
            }
        }
    }

    /// <summary>
    /// MsiInstallerクラスの例外クラス
    /// </summary>
    class MsiInstallerException : Exception
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="message">エラーを説明するメッセージ。 </param>
        public MsiInstallerException(string message) : base(message) { }
    }
}
