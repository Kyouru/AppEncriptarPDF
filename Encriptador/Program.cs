using System;
using System.Text;
using System.IO;
using System.Configuration;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace Encriptador
{
    class Program
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        static string pdforigen;
        static string pdfdestino;
        static int intervalo;
        static int limite;

        static void Main(string[] args)
        {
            const int SW_HIDE = 0;
            //const int SW_SHOW = 5;
            var handle = GetConsoleWindow();

            Process currentProcess = Process.GetCurrentProcess();
            Console.Title = "Encriptar PDF";
            pdforigen = ConfigurationManager.AppSettings["ruta_origen"].ToString();
            pdfdestino = ConfigurationManager.AppSettings["ruta_destino"].ToString();
            intervalo = Int32.Parse(ConfigurationManager.AppSettings["intervalo"].ToString());
            limite = Int32.Parse(ConfigurationManager.AppSettings["limite"].ToString());

            string manual = "false";
            int offset = 0;

            while (args.Length - offset > 0)
            {
                if (args[0 + offset] == "-help" || args[0 + offset] == "--help" || args[0 + offset] == "-hide" || args[0 + offset] == "-killall" || args[0 + offset] == "-manual" || args[0 + offset] == "-intervalo" || args[0 + offset] == "-limite" || args[0 + offset] == "-rutaorigen" || args[0 + offset] == "-rutadestino")
                {
                    if (args[0 + offset] == "-help" || args[0 + offset] == "--help")
                    {
                        mostrarAyuda();
                        Environment.Exit(0);
                    }
                    else if (args[0 + offset] == "-hide")
                    {
                        ShowWindow(handle, SW_HIDE);
                        //ShowWindow(handle, SW_SHOW);
                        offset++;
                    }
                    else if (args[0 + offset] == "-killall")
                    {
                        foreach (var process in Process.GetProcessesByName("Encriptador"))
                        {
                            if (process.Id != currentProcess.Id)
                            {
                                process.Kill();
                                Console.WriteLine(" Encriptador con PID " + process.Id + " cerrado.");
                            }
                        }
                        Environment.Exit(0);
                    }
                    else if (args[0 + offset] == "-manual")
                    {
                        manual = "true";
                        offset++;
                    }
                    else if (args[0 + offset] == "-intervalo")
                    {
                        try
                        {
                            intervalo = Convert.ToInt32(args[1 + offset]);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("  Error. " + ex.Message);
                            mostrarAyuda();
                            Environment.Exit(0);
                        }
                        offset += 2;
                    }
                    else if (args[0 + offset] == "-limite")
                    {
                        try
                        {
                            limite = Convert.ToInt32(args[1 + offset]);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("  Error. " + ex.Message);
                            mostrarAyuda();
                            Environment.Exit(0);
                        }
                        offset += 2;
                    }
                    else if (args[0 + offset] == "-rutaorigen")
                    {
                        if (Directory.Exists(args[1 + offset]))
                        {
                            pdforigen = args[1 + offset];
                        }
                        else
                        {
                            Console.WriteLine("  Error. Ruta Origen no existe");
                            mostrarAyuda();
                            Environment.Exit(0);
                        }
                        offset += 2;
                    }
                    else if (args[0 + offset] == "-rutadestino")
                    {
                        if (Directory.Exists(args[1 + offset]))
                        {
                            pdfdestino = args[1 + offset];
                        }
                        else
                        {
                            Console.WriteLine("  Error. Ruta Destino no existe");
                            mostrarAyuda();
                            Environment.Exit(0);
                        }
                        offset += 2;
                    }
                }
                else
                {
                    Console.WriteLine("  Error. Revisar parametros");
                    mostrarAyuda();
                    Environment.Exit(0);
                }
            }

            string rutainput2 = SelectFromWhere("SELECT TBLDETALLE FROM SYST900 S WHERE TBLCODTAB = 50 AND TBLESTADO = '1' AND tblcodarg IN (17)", false) + "\\";

            Console.WriteLine(rutainput2);
            Console.WriteLine(pdforigen);

            FileSystemWatcher fileSystemWatcher = new FileSystemWatcher();
            FileSystemWatcher fileSystemWatcherExcelPCT = new FileSystemWatcher();

            if (manual != "true")
            {

                fileSystemWatcher.Path = pdforigen;
                fileSystemWatcher.Created += OnCreated;
                fileSystemWatcher.EnableRaisingEvents = true;
                //fileSystemWatcher.IncludeSubdirectories = false;

                fileSystemWatcherExcelPCT.Path = rutainput2;
                fileSystemWatcherExcelPCT.Created += ExecuteExcelPCT;
                fileSystemWatcherExcelPCT.EnableRaisingEvents = true;
                //fileSystemWatcherExcelPCT.IncludeSubdirectories = false;

                Console.Read();
            }
            else
            {
                string[] dirs = Directory.GetFiles(pdforigen);
                if (dirs.Length > 0)
                {
                    Console.WriteLine("\nRuta Password PDF: " + pdforigen);
                    Console.WriteLine("Se encontró " + dirs.Length + " archivos:");

                    foreach (string dir in dirs)
                    {
                        //Console.WriteLine("   > " + Path.GetFileName(dir));
                        FileSystemEventArgs fsea = new FileSystemEventArgs(WatcherChangeTypes.Created, pdforigen, Path.GetFileName(dir));
                        OnCreated(fileSystemWatcher, fsea);
                    }
                }

                dirs = Directory.GetFiles(rutainput2);
                if (dirs.Length > 0)
                {
                    Console.WriteLine("\nRuta Excel PCT: " + rutainput2);
                    Console.WriteLine("Se encontró " + dirs.Length + " archivos:");

                    foreach (string dir in dirs)
                    {
                        //Console.WriteLine("   > " + Path.GetFileName(dir));
                        FileSystemEventArgs fsea = new FileSystemEventArgs(WatcherChangeTypes.Created, rutainput2, Path.GetFileName(dir));
                        ExecuteExcelPCT(fileSystemWatcherExcelPCT, fsea);
                    }
                }
            }
        }

        static void mostrarAyuda()
        {
            Console.WriteLine(" Ayuda:");
            Console.WriteLine("  -hide: Oculta la consola");
            Console.WriteLine("  -killall: Termina todos los procesos en ejecucion de nombre Encriptador.exe");
            Console.WriteLine("  -manual: Para ejecutar a demanda");
            Console.WriteLine("  -intervalo <tiempo ms>: Invervalo entre los reintentos en caso .pdf bloqueado");
            Console.WriteLine("  -limite <tiempo ms>: Limite total de los reintentos en caso .pdf bloqueado");
            Console.WriteLine("  -rutaorigen <ruta>: Procesa los archivos en la <ruta>");
            Console.WriteLine("  -rutadestino <ruta>: Deja los .pdf con clave en la <ruta>");
        }

        static void OnCreated(object sender, FileSystemEventArgs e)
        {

            FileInfo fi = new FileInfo(pdforigen + e.Name);

            int cont = 0;
            //validar que el archivo no este en proceso de copia o este en uso por otra aplicacion
            while (IsFileLocked(fi))
            {
                Thread.Sleep(Int32.Parse(ConfigurationManager.AppSettings["intervalo"].ToString()));
                cont++;
                if (cont >= Int32.Parse(ConfigurationManager.AppSettings["limite"].ToString())/Int32.Parse(ConfigurationManager.AppSettings["intervalo"].ToString()))
                {
                    break;
                }
            }

            try
            {
                Document document = new Document();
                PdfReader reader = new PdfReader(pdforigen + e.Name);
                PdfStamper stamper = new PdfStamper(reader, new FileStream(pdfdestino + e.Name, FileMode.Create));
                stamper.SetEncryption(Encoding.ASCII.GetBytes("Cts2019C00Pac"),
                                        Encoding.ASCII.GetBytes(Obtener_DNI(e.Name.ToString().Substring(0, 7))),
                                        PdfWriter.ALLOW_PRINTING, PdfWriter.ENCRYPTION_AES_128
                               | PdfWriter.DO_NOT_ENCRYPT_METADATA);
                stamper.Close();
                reader.Close();
                File.Delete(pdforigen + e.Name);
                Console.WriteLine(pdforigen + e.Name + " >> " + pdfdestino + e.Name);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error. " + ex.Message + "\n");
            }
        }
        public static string Obtener_DNI(string vCodSocio)
        {
            string connectionString = ConfigurationManager.AppSettings["OracleCnx"].ToString();
            string queryString =
                "SELECT PN.NUMERODOCUMENTOID FROM PERSONA P, PERSONANATURAL PN WHERE P.CODIGOPERSONA=PN.CODIGOPERSONA AND P.CIP=" + vCodSocio.ToString();
            string vDNI = "";
            using (OracleConnection connection =
                   new OracleConnection(connectionString))
            {
                OracleCommand command = connection.CreateCommand();
                command.CommandText = queryString;
                try
                {
                    connection.Open();
                    OracleDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        vDNI = reader[0].ToString();
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            return vDNI;
        }
        static void ExecuteExcelPCT(object sender, FileSystemEventArgs e)
        {
            string rutainput = SelectFromWhere("SELECT TBLDETALLE FROM SYST900 S WHERE TBLCODTAB = 50 AND TBLESTADO = '1' AND tblcodarg IN (17)", false) + "\\";
            string rutaoutput = SelectFromWhere("SELECT TBLDETALLE FROM SYST900 S WHERE TBLCODTAB = 50 AND TBLESTADO = '1' AND tblcodarg IN (18)", false) + "\\";
            string rutarejected = SelectFromWhere("SELECT TBLDETALLE FROM SYST900 S WHERE TBLCODTAB = 50 AND TBLESTADO = '1' AND tblcodarg IN (19)", false) + "\\";
            string file = e.Name;

            string varchivovalido = SelectFromWhere("SELECT substr(descripcionarchivo,1,instr(descripcionarchivo,'.',1,1)-1) FROM (SELECT 'ENC.XLS' AS descripcionarchivo FROM DUAL UNION ALL SELECT 'DET.XLS' AS descripcionarchivo FROM DUAL) WHERE substr(descripcionarchivo,1,instr(descripcionarchivo,'.',1,1)-1) IS NOT NULL AND (UPPER('" + file + "') LIKE '%'||substr(descripcionarchivo,1,instr(descripcionarchivo,'.',1,1)-1)||'%') GROUP BY descripcionarchivo ORDER BY descripcionarchivo", true);
            string vextensionvalida = SelectFromWhere("SELECT TRIM(substr(descripcionarchivo, instr(descripcionarchivo, '.', 1, 1) + 1, 4)) FROM (SELECT 'ENC.XLS' AS descripcionarchivo FROM DUAL UNION ALL SELECT 'DET.XLS' AS descripcionarchivo FROM DUAL) WHERE substr(descripcionarchivo,1,instr(descripcionarchivo,'.',1,1)-1) IS NOT NULL AND (UPPER('" + file + "') LIKE '%'||TRIM(substr(descripcionarchivo,instr(descripcionarchivo,'.',1,1)+1,4))||'%') AND ROWNUM = 1 GROUP BY descripcionarchivo ORDER BY descripcionarchivo", true);

            if (!(String.IsNullOrEmpty(varchivovalido)) && !(String.IsNullOrEmpty(vextensionvalida)))
            {
                Console.WriteLine(rutainput + file);
                string xlsFilePath = Path.Combine(rutainput, file);
                read_filePCT(xlsFilePath, varchivovalido);
                readed_filePCT(rutainput, rutaoutput, rutarejected);
            }
            else
            {
                string fecharechaza = SelectFromWhere("SELECT TO_CHAR(SYSDATE, 'dd-mm-YYYY HH24MISS') FROM DUAL", true);
                if (File.Exists(rutainput + file))
                {
                   File.Move(rutainput + file, rutarejected + fecharechaza + "_" + file);
                }
            }
        }

        public static void readed_filePCT(string prmtrutainput, string prmtrutaoutput, string prmtrutarejected)
        {
            string connectionString = ConfigurationManager.AppSettings["OracleCnx"].ToString();
            string queryString = "SELECT DISTINCT acbtmp.ID_ARCHIVO, SUBSTR(acbtmp.NOMBREARCHIVO,(INSTR(acbtmp.NOMBREARCHIVO,'\\',-1)+1)) FILEIN,(SUBSTR(SUBSTR(acbtmp.NOMBREARCHIVO, (INSTR(acbtmp.NOMBREARCHIVO, '\\',-1)+1)),1,INSTR(SUBSTR(acbtmp.NOMBREARCHIVO,(INSTR(acbtmp.NOMBREARCHIVO,'\\',-1)+1)),'.',1,1)-1)) ||(CASE MOD(acbtmp.ESTADO, 2) WHEN 1 THEN '_APROBADO' ELSE '_RECHAZADO' END) || '_' || (TO_CHAR(SYSDATE, 'dd-mm-YYYY HH24MISS') || '.' ||SUBSTR(acbtmp.NOMBREARCHIVO,(INSTR(acbtmp.NOMBREARCHIVO, '.', -1, 1) + 1), length(acbtmp.NOMBREARCHIVO))) As FileOut FROM CRONOGRAMACOFIDETMP acbtmp";

            int vid_archivo = 0;
            int vrechazado = 0;
            string vfileinput = "";
            string vfileoutput = "";

            using (OracleConnection connection =
                   new OracleConnection(connectionString))
            {
                OracleCommand command = connection.CreateCommand();
                command.CommandText = queryString;
                try
                {
                    bool exists = Directory.Exists(prmtrutaoutput);

                    if (!exists)
                        Directory.CreateDirectory(prmtrutaoutput);

                    connection.Open();
                    OracleDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        vid_archivo = Convert.ToInt32(reader[0].ToString());
                        vfileinput = reader[1].ToString();
                        vfileoutput = reader[2].ToString();
                        vrechazado = vfileoutput.IndexOf("RECHAZADO");
                        string sourceFile = Path.Combine(prmtrutainput, vfileinput);
                        string destFile = Path.Combine(prmtrutaoutput, vfileoutput);
                        string destFileR = Path.Combine(prmtrutarejected, vfileoutput);

                        if (File.Exists(sourceFile))
                        {
                            if(vrechazado < 0)
                            {
                                File.Move(sourceFile, destFile);
                            }
                            else
                            {
                                File.Move(sourceFile, destFileR);
                            }                            
                        }                        
                    }
                    reader.Close();
                    InsUpdDel_Oracle("DELETE FROM CRONOGRAMACOFIDETMP WHERE ID_ARCHIVO = " + vid_archivo);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
        public static void read_filePCT(string xlsFilePath, string ArchivoValido)
        {
            if (!File.Exists(xlsFilePath))
                return;

            FileInfo fi = new FileInfo(xlsFilePath);
            long filesize = fi.Length;

            Application xlApp;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            Range range;
            var misValue = Type.Missing;//System.Reflection.Missing.Value;

            // abrir el documento 
            xlApp = new Application();
            xlWorkBook = xlApp.Workbooks.Open(xlsFilePath, misValue, misValue,
            misValue, misValue, misValue, misValue, misValue, misValue,
            misValue, misValue, misValue, misValue, misValue, misValue);

            // seleccion de la hoja de calculo
            // get_item() devuelve object y numera las hojas a partir de 1
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            // seleccion rango activo
            range = xlWorkSheet.UsedRange;
            range.Columns.AutoFit();

            // Variables de Campos
            string vCampo_A, vCampo_B, vCampo_C, vCampo_D, vCampo_E, vCampo_F, vCampo_G, vCampo_H, vCampo_I, vCampo_J, vCampo_K, vCampo_L;
            string queryinsert, queryvalues;

            // leer las celdas
            int nFilaNada = 0, nFilaAlgo = 0;
            int rows = range.Rows.Count;
            int cols = range.Columns.Count;
            int colval = 0;

            if (cols > 12) cols = 12;

            int vId_Archivo = Convert.ToInt32(SelectFromWhere("SELECT(NVL(MAX(P.ID_ARCHIVO), 0) + 1) AS MAX_ID FROM CRONOGRAMACOFIDETMP P", false));

            DateTime hoy = DateTime.Now;

            for (int row = 1; row <= rows; row++)
            {
                vCampo_A = ""; vCampo_B = ""; vCampo_C = ""; vCampo_D = "";
                vCampo_E = ""; vCampo_F = ""; vCampo_G = ""; vCampo_H = "";
                vCampo_I = ""; vCampo_J = ""; vCampo_K = ""; vCampo_L = "";

                for (int col = 1; col <= cols; col++)
                {
                    // lectura como cadena
                    var cellText = xlWorkSheet.Cells[row, col].Text;
                    cellText = Convert.ToString(cellText);
                    cellText = cellText.Replace("'", ""); // Comillas simples no pueden pasar en el Texto

                    switch (col)
                    {
                        case 1: vCampo_A = cellText; break;
                        case 2: vCampo_B = cellText; break;
                        case 3: vCampo_C = cellText; break;
                        case 4: vCampo_D = cellText; break;
                        case 5: vCampo_E = cellText; break;
                        case 6: vCampo_F = cellText; break;
                        case 7: vCampo_G = cellText; break;
                        case 8: vCampo_H = cellText; break;
                        case 9: vCampo_I = cellText; break;
                        case 10: vCampo_J = cellText; break;
                        case 11: vCampo_K = cellText; break;
                        case 12: vCampo_L = cellText; break;
                    }
                }

                if (row == 1)
                {
                    if (!(String.IsNullOrEmpty(vCampo_A.Trim())) && !(String.IsNullOrEmpty(vCampo_B.Trim())) && !(String.IsNullOrEmpty(vCampo_C.Trim())) && !(String.IsNullOrEmpty(vCampo_D.Trim())) && !(String.IsNullOrEmpty(vCampo_E.Trim())) && !(String.IsNullOrEmpty(vCampo_F.Trim())) && !(String.IsNullOrEmpty(vCampo_G.Trim())) && !(String.IsNullOrEmpty(vCampo_H.Trim())) && !(String.IsNullOrEmpty(vCampo_I.Trim())) && !(String.IsNullOrEmpty(vCampo_J.Trim())) && !(String.IsNullOrEmpty(vCampo_K.Trim())) && !(String.IsNullOrEmpty(vCampo_L.Trim())))
                    {
                        colval = 0;
                    }
                    else
                    {
                        colval = 1;
                        queryinsert = "INSERT INTO CRONOGRAMACOFIDETMP (CONTRATO, SECUENCIA, CUOTA, FECHAVENCIMIENTO, NDIAS, MONEDA, PRINCIPAL, INTERES, COMISION, MONTOCOBRAR, PRINCIPALVENCER, CAPITALIZACION, ID_ARCHIVO, NOMBREARCHIVO, TAMANOARCHIVO, ID_FILAS, ESTADO, ARCHIVOVALIDO) ";
                        queryvalues = "VALUES ('" + vCampo_A + "', '" + vCampo_B + "', '" + vCampo_C + "', '" + vCampo_D + "', '" + vCampo_E + "', '" + vCampo_F + "', '" + vCampo_G + "', '" + vCampo_H + "', '" + vCampo_I + "', '" + vCampo_J + "', '" + vCampo_K + "', '" + vCampo_L + "', " + vId_Archivo + ", '" + xlsFilePath + "', " + filesize + ", " + row + ", " + 0 /* 0 para Estado Carga Inicial */ + ", '" + ArchivoValido + "')";
                        InsUpdDel_Oracle(queryinsert + queryvalues);
                    }
                }
                else
                {
                    if (colval == 0)
                    {
                        if (String.IsNullOrEmpty(vCampo_A.Trim()) && String.IsNullOrEmpty(vCampo_B.Trim()) && String.IsNullOrEmpty(vCampo_C.Trim()) && String.IsNullOrEmpty(vCampo_D.Trim()) && String.IsNullOrEmpty(vCampo_E.Trim()) && String.IsNullOrEmpty(vCampo_F.Trim()) && String.IsNullOrEmpty(vCampo_G.Trim()) && String.IsNullOrEmpty(vCampo_H.Trim()) && String.IsNullOrEmpty(vCampo_I.Trim()) && String.IsNullOrEmpty(vCampo_J.Trim()) && String.IsNullOrEmpty(vCampo_K.Trim()) && String.IsNullOrEmpty(vCampo_L.Trim()))
                        {
                            nFilaNada++;
                        }
                        else
                        {
                            nFilaAlgo = row;
                            nFilaNada = 0;
                        }

                        if (nFilaNada > 10)
                            rows = row++;
                        queryinsert = "INSERT INTO CRONOGRAMACOFIDETMP (CONTRATO, SECUENCIA, CUOTA, FECHAVENCIMIENTO, NDIAS, MONEDA, PRINCIPAL, INTERES, COMISION, MONTOCOBRAR, PRINCIPALVENCER, CAPITALIZACION, ID_ARCHIVO, NOMBREARCHIVO, TAMANOARCHIVO, ID_FILAS, ESTADO, ARCHIVOVALIDO) ";
                        queryvalues = "VALUES ('" + vCampo_A + "', '" + vCampo_B + "', '" + vCampo_C + "', '" + vCampo_D + "', '" + vCampo_E + "', '" + vCampo_F + "', '" + vCampo_G + "', '" + vCampo_H + "', '" + vCampo_I + "', '" + vCampo_J + "', '" + vCampo_K + "', '" + vCampo_L + "', " + vId_Archivo + ", '" + xlsFilePath + "', " + filesize + ", " + row + ", " + 0 /* 0 para Estado Carga Inicial */ + ", '" + ArchivoValido + "')";
                        InsUpdDel_Oracle(queryinsert + queryvalues);
                    }
                }
            }

            if (colval == 1)
            {
                int vAproRecha = 2;
                InsUpdDel_Oracle("UPDATE CRONOGRAMACOFIDETMP SET ESTADO = " + vAproRecha + " WHERE ID_ARCHIVO = " + vId_Archivo);
            }
            else
            {
                InsUpdDel_Oracle("DELETE FROM CRONOGRAMACOFIDETMP WHERE ID_ARCHIVO = " + vId_Archivo + " AND ID_FILAS > " + nFilaAlgo);

                int vAproRecha = Convert.ToInt32(Function_Procedure_Oracle(1, "CRONOGRAMA_TECHO_PROPIO.F_CARGACRONOGRAMALINEAPTP", "PIid_archivo", vId_Archivo, "", -1, "", -1));
                InsUpdDel_Oracle("UPDATE CRONOGRAMACOFIDETMP SET ESTADO = " + vAproRecha + " WHERE ID_ARCHIVO = " + vId_Archivo);
            }

            // cerrar
            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();

            // liberar
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
 
        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to release the object(object:{0})\n" + ex.Message, obj.ToString());
            }
            finally
            {
                obj = null;
                GC.Collect();
            }
        }
        public static string SelectFromWhere(string executequery, bool eslike)
        {
            string connectionString = ConfigurationManager.AppSettings["OracleCnx"].ToString();
            string queryString = executequery.ToString();

            string vDATO = "";

            using (OracleConnection connection =
                   new OracleConnection(connectionString))
            {
                OracleCommand command = connection.CreateCommand();
                command.CommandText = queryString;
                try
                {
                    connection.Open();
                    OracleDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        vDATO = reader[0].ToString();
                        if (eslike) break;
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            return vDATO;
        }
        public static void InsUpdDel_Oracle(string executequery)
        {
            string connectionString = ConfigurationManager.AppSettings["OracleCnx"].ToString();
            string queryString = executequery.ToString();
            string queryCommit = "COMMIT";

            using (OracleConnection connection =
                   new OracleConnection(connectionString))
            {
                OracleCommand command = connection.CreateCommand();                
                try
                {
                    connection.Open();
                    command.CommandType = CommandType.Text;
                    command.CommandText = queryString;
                    command.ExecuteNonQuery();
                    command.CommandText = queryCommit;
                    command.ExecuteNonQuery();

                    connection.Close();
                    command.Dispose();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
        public static string Function_Procedure_Oracle(int tipofunpro /* tipofunpro: 1 para function, 2 para procedure  */, string executequery, string nomprmt1, int prmt1, string nomprmt2, int prmt2, string nomprmt3, int prmt3)
        {
            string connectionString = ConfigurationManager.AppSettings["OracleCnx"].ToString();
            string queryString = executequery.ToString();

            string vDATO = "";

            using (OracleConnection connection =
                   new OracleConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    OracleCommand command = connection.CreateCommand();
                    command.CommandText = queryString;
                    command.CommandType = CommandType.StoredProcedure;
                    if (prmt1 >= 0)
                    {
                        command.Parameters.Add(new OracleParameter(nomprmt1, OracleDbType.Int32)).Value = prmt1;
                        command.Parameters[nomprmt1].Direction = ParameterDirection.Input;
                    }
                    if (prmt2 >= 0)
                    {
                        command.Parameters.Add(new OracleParameter(nomprmt2, OracleDbType.Int32)).Value = prmt2;
                        command.Parameters[nomprmt2].Direction = ParameterDirection.Input;
                    }
                    if (prmt3 >= 0)
                    {
                        command.Parameters.Add(new OracleParameter(nomprmt3, OracleDbType.Int32)).Value = prmt3;
                        command.Parameters[nomprmt3].Direction = ParameterDirection.Input;
                    }

                    if (tipofunpro == 1)
                    {
                        command.Parameters.Add("retorno", OracleDbType.Int32);
                        command.Parameters["retorno"].Direction = ParameterDirection.ReturnValue;
                    }                    

                    command.ExecuteNonQuery();

                    if (tipofunpro == 1)
                        vDATO = Convert.ToString(command.Parameters["retorno"].Value);

                    connection.Close();
                    command.Dispose();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            return vDATO;
        }

        static private bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                Console.WriteLine(file.FullName + " Bloqueado");
                return true;
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                }
            }
            return false;
        }
    }    
}
