using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO.Ports;

namespace SMEtiquetasBrother
{
    public class clGlobal
    {
        [DllImport("Kernel32.dll")]
        public static extern bool QueryPerformanceCounter(out long lpPerformanceCount);
        [DllImport("Kernel32.dll")]
        public static extern bool QueryPerformanceFrequency(out long lpFrequency);

        public const string EnderecoIPPRODSERV = "192.168.25.3";
        public const string EnderecoLocalPRODSERV = "PRODSERV";
        public const string EnderecoLocalPRODSERVdominio = "PRODSERV.mgers2016.com.br";
        public const string EnderecoRemotoPRODSERV = "starmeasure.ddns.net";
        public const string UsuarioMySQL = "SMEtiquetas";
        public const string SenhaMySQL = "SMEtiquetas@xyzmge123";

        public static string HostName = "";
        public static string EnderecoServidor = "";

        public static string MessageText = "";
        public static string MessageWarning = "";

        public static bool bTerminouPing = false;
        public static bool bRespostaPing = false;

        public static int QtdEtiquetasIndividuais = 0;
        public static int QtdEtiquetasColetivas = 0;
    }
}