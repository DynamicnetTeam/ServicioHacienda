using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using System.Net;
using System.Net.Mail;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace ServicioHacienda
{
    public partial class ServicioHacienda : ServiceBase
    {
        MySqlConnection _cnn;
        Boolean _EnProceso = false;

        #region Eventos
        public ServicioHacienda()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            string _Connectionstring;
            _Connectionstring = "server=localhost;port=3306;user id=root;password=admin123;database=db_electronic;pooling=false;";
            DB_Conectar(_Connectionstring);
        }

        protected override void OnStop()
        {
            DB_Desconectar();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (_EnProceso == false)
            {
                _EnProceso = true;
                Procesar();
                _EnProceso = false;
            }
        }
        #endregion

        #region DB_Conectar
        /// <summary>
        /// Ejecuta la conexion con la base de datos
        /// </summary>
        private void DB_Conectar(string pConnetionString)
        {
            _cnn = new MySqlConnection(pConnetionString);

            try
            {
                _cnn.Open();
            }
            catch 
            {

            }
        }
        #endregion 

        #region DB_Desconectar
        /// <summary>
        /// Se desconecta de la base de datos
        /// </summary>
        private void DB_Desconectar()
        {
            _cnn.Close();
        }
        #endregion 

        #region SendEmail
        public void SendEmail(DataTable dt1, string pReceptorEmail, string pXMLFirmado, string pPDF)
        {
            try
            {
                string mEmailEmail = string.Empty;
                string mEmailServidor = string.Empty;
                string mEmailUsusrio = string.Empty;
                string mEmailClave = string.Empty;
                int mEmailPuerto = 0;
                string attachmentFilename;

                var fromAddress = new MailAddress("fernandosolis5@gmail.com", "Dynamic");
                var toAddress = new MailAddress("fernandosolis5@gmail.com", "Dynamic");
                const string fromPassword = "heroeleyenda555";// "oblmdzxaqrwprwxp";
                const string subject = "Subject";
                const string body = "Body";

                if (dt1.Rows.Count > 0)
                {
                  DataRow drInfo = dt1.Rows[0];

                  if (drInfo["Email"] != System.DBNull.Value)
                        mEmailEmail = drInfo["Email"].ToString();
                  else
                        mEmailEmail = string.Empty;
                  if (drInfo["Servidor"] != System.DBNull.Value)
                        mEmailServidor = drInfo["Servidor"].ToString();
                  else
                        mEmailServidor = string.Empty;
                  if (drInfo["Ususrio"] != System.DBNull.Value)
                        mEmailUsusrio = drInfo["Ususrio"].ToString();
                  else
                        mEmailUsusrio = string.Empty;
                  if (drInfo["Clave"] != System.DBNull.Value)
                        mEmailClave = drInfo["Clave"].ToString();
                  else
                        mEmailClave = string.Empty;
                  if (drInfo["Puerto"] != System.DBNull.Value)
                        mEmailPuerto = (int)drInfo["Puerto"];
                  else
                        mEmailPuerto = 0;

                    MailMessage mail = new MailMessage();
                    SmtpClient SmtpServer = new SmtpClient(mEmailServidor);
                    mail.From = new MailAddress(mEmailEmail);
                    mail.To.Add(pReceptorEmail);
                    mail.Subject = "Factura electronica: " + pXMLFirmado;
                    mail.Body = "Estimado(a): Adjunto a este correo encontrará un Comprobante Electrónico en formato XML y su correspondiente representación en formato PDF.";

                    System.Net.Mail.Attachment attachment1;
                    System.Net.Mail.Attachment attachment2;
                    if (pXMLFirmado != string.Empty)
                    {
                        attachment1 = new System.Net.Mail.Attachment(pXMLFirmado);
                        mail.Attachments.Add(attachment1);
                    }
                    if (pPDF != string.Empty)
                    {
                        attachment2 = new System.Net.Mail.Attachment(pPDF);
                        mail.Attachments.Add(attachment2);
                    }
                                                                
                    SmtpServer.Port = mEmailPuerto;
                    SmtpServer.UseDefaultCredentials = false;
                    SmtpServer.Credentials = new System.Net.NetworkCredential(mEmailUsusrio, mEmailClave);
                    SmtpServer.EnableSsl = true;

                    SmtpServer.Send(mail);

                    /*
                    var smtp = new SmtpClient
                    {
                        Host = mEmailServidor,
                    Port = mEmailPuerto,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(mEmailUsusrio, mEmailClave)                    
                    };
                using (var message = new MailMessage(mEmailUsusrio, pReceptorEmail)
                {
                    Subject = subject,
                    Body = body
                    
                })
                {
                    smtp.Send(message);
                }*/
                }
            }
            catch
            { }
        }
        #endregion

        #region Get_HaciendaDocuments
        /// <summary>
        /// Consulta los documentos no procesados para enviarlos a GTI
        /// </summary>
        private void Get_HaciendaDocuments(DataTable dt1)
        {
            MySqlCommand command;
            MySqlDataAdapter sda;
            StringBuilder sql = new StringBuilder();
            sql.AppendLine("Select * From tb_dcumentosencabezado ");
            sql.AppendLine("Where (IdEstadoHacienda = 0) ");
            sql.AppendLine("Or (CodigoEmail = 0)");
            try
            {
                command = new MySqlCommand(sql.ToString(), _cnn);
                command.ExecuteNonQuery();
                sda = new MySqlDataAdapter(command);
                sda.Fill(dt1);
                command.Dispose();
            }
            catch (Exception ex)
            {

            }
        }
        #endregion

        #region Get_EmpresaEmail        
        private void Get_EmpresaEmail(DataTable dt1, int pIdEmpresa)
        {
            MySqlCommand command;
            MySqlDataAdapter sda;
            StringBuilder sql = new StringBuilder();
            sql.AppendLine("Select * From tb_emisoremail ");
            sql.AppendLine("Where IdEmpresa = @IdEmpresa");
            try
            {
                command = new MySqlCommand(sql.ToString(), _cnn);
                command.Parameters.Add("IdEmpresa", MySqlDbType.Int32).Value = pIdEmpresa;
                command.ExecuteNonQuery();
                sda = new MySqlDataAdapter(command);
                sda.Fill(dt1);
                command.Dispose();
            }
            catch (Exception ex)
            {

            }
        }
        #endregion

        #region Get_EmpresaIdPlsntilla       
        private int Get_EmpresaIdPlsntilla(int pIdEmpresa)
        {
            int mResultado = 0;
            MySqlCommand command;
            DataTable dt1 = new DataTable();
            MySqlDataAdapter sda;
            StringBuilder sql = new StringBuilder();
            sql.AppendLine("Select IdPlantilla From tb_empresas ");
            sql.AppendLine("Where Id = @IdEmpresa");
            try
            {
                command = new MySqlCommand(sql.ToString(), _cnn);
                command.Parameters.Add("IdEmpresa", MySqlDbType.Int32).Value = pIdEmpresa;
                command.ExecuteNonQuery();
                sda = new MySqlDataAdapter(command);
                sda.Fill(dt1);
                if (dt1.Rows.Count > 0)
                    mResultado = (int)dt1.Rows[0][0];
                command.Dispose();
            }
            catch (Exception ex)
            {

            }

            return mResultado;
        }
        #endregion

        #region Get_EmpresaToken        
        private void Get_EmpresaToken(int pIdEmpresa, ref string pTokenUsuario, ref string pTokenClave) 
        {
            pTokenUsuario = string.Empty;
            pTokenClave = string.Empty;
            MySqlCommand command;
            DataTable dt1 = new DataTable();
            MySqlDataAdapter sda;
            StringBuilder sql = new StringBuilder();
            sql.AppendLine("Select TokenUsuario,TokenClave From tb_empresas ");
            sql.AppendLine("Where Id = @IdEmpresa");
            try
            {
                command = new MySqlCommand(sql.ToString(), _cnn);
                command.Parameters.Add("IdEmpresa", MySqlDbType.Int32).Value = pIdEmpresa;
                command.ExecuteNonQuery();
                sda = new MySqlDataAdapter(command);
                sda.Fill(dt1);
                if (dt1.Rows.Count > 0)
                {
                    pTokenUsuario = dt1.Rows[0][0].ToString();
                    pTokenClave = dt1.Rows[0][1].ToString();
                }
                command.Dispose();
            }
            catch (Exception ex)
            {

            }
        }
        #endregion

        #region Get_Plantilla       
        private string Get_Plantilla(int pId)
        {
            string mResultado = string.Empty;
            MySqlCommand command;
            DataTable dt1 = new DataTable();
            MySqlDataAdapter sda;
            StringBuilder sql = new StringBuilder();
            sql.AppendLine("Select Plantilla From tb_plantillas ");
            sql.AppendLine("Where Id = @Id");
            try
            {
                command = new MySqlCommand(sql.ToString(), _cnn);
                command.Parameters.Add("Id", MySqlDbType.Int32).Value = pId;
                command.ExecuteNonQuery();
                sda = new MySqlDataAdapter(command);
                sda.Fill(dt1);
                if (dt1.Rows.Count > 0)
                    mResultado = dt1.Rows[0][0].ToString();
                command.Dispose();
            }
            catch (Exception ex)
            {

            }
            return mResultado;
        }
        #endregion

        #region Actualizar_HaciendaEmailCodigo
        private void Actualizar_HaciendaEmailCodigo(int pCodigo, int pId)
        {
            MySqlCommand command;
            StringBuilder sql = new StringBuilder();
            sql.AppendLine("Update tb_dcumentosencabezado ");
            sql.AppendLine("Set ");
            sql.AppendLine("CodigoEmail = @CodigoEmail");            
            sql.AppendLine("Where Id = @Id");
            try
            {
                command = new MySqlCommand(sql.ToString(), _cnn);
                command.Parameters.Add("CodigoEmail", MySqlDbType.Int32).Value = pCodigo;
                command.Parameters.Add("Id", MySqlDbType.Int32).Value = pId;
                command.ExecuteNonQuery();
                command.Dispose();
            }
            catch (Exception ex)
            {

            }
        }
        #endregion

        #region Actualizar_Hacienda
        private void Actualizar_Hacienda(int pIdEstadoHacienda, int pId, string pMensaje, string pDetalleHacienda)
        {
            MySqlCommand command;
            StringBuilder sql = new StringBuilder();
            sql.AppendLine("Update tb_dcumentosencabezado ");
            sql.AppendLine("Set ");
            sql.AppendLine("IdEstadoHacienda = @IdEstadoHacienda,");
            sql.AppendLine("Mensaje = @Mensaje,");
            sql.AppendLine("DetalleHacienda = @DetalleHacienda");
            sql.AppendLine("Where Id = @Id");
            try
            {
                command = new MySqlCommand(sql.ToString(), _cnn);
                command.Parameters.Add("IdEstadoHacienda", MySqlDbType.Int32).Value = pIdEstadoHacienda;
                command.Parameters.Add("Id", MySqlDbType.Int32).Value = pId;
                command.Parameters.Add("Mensaje", MySqlDbType.String).Value = pMensaje;
                command.Parameters.Add("DetalleHacienda", MySqlDbType.String).Value = pDetalleHacienda;
                command.ExecuteNonQuery();
                command.Dispose();
            }
            catch (Exception ex)
            {

            }
        }
        #endregion

        #region Procesar
        public void Procesar()
        {
            try
            {
                DataTable dtDocuments = new DataTable();
                DataTable dtEmail = new DataTable();

                string mLinea = string.Empty;
                int IdPlsntilla = 0;
                string Plsntilla = string.Empty;
                int mId = 0;
                int mIdEmpresa = 0;
                string mReceptorCorreos = string.Empty;
                string mClave = string.Empty;
                string mXMLConFirma = string.Empty;
                int mCodigoEmail = 0;
                int mIdEstadoHacienda = 0;

                Get_HaciendaDocuments(dtDocuments);
                foreach (DataRow rDocuments in dtDocuments.Rows)
                {
                    if (rDocuments["Id"] != System.DBNull.Value)
                        mId = (int)rDocuments["Id"];
                    else
                        mId = 0;
                    if (rDocuments["IdEmpresa"] != System.DBNull.Value)
                        mIdEmpresa = (int)rDocuments["IdEmpresa"];
                    else
                        mIdEmpresa = 0;
                    if (rDocuments["ReceptorCorreos"] != System.DBNull.Value)
                        mReceptorCorreos = rDocuments["ReceptorCorreos"].ToString();
                    else
                        mReceptorCorreos = string.Empty;
                    if (rDocuments["Clave"] != System.DBNull.Value)
                        mClave = rDocuments["Clave"].ToString();
                    else
                        mClave = string.Empty;
                    if (rDocuments["XMLConFirma"] != System.DBNull.Value)
                        mXMLConFirma = rDocuments["XMLConFirma"].ToString();
                    else
                        mXMLConFirma = string.Empty; ;
                    if (rDocuments["CodigoEmail"] != System.DBNull.Value)
                        mCodigoEmail = (int)rDocuments["CodigoEmail"];
                    else
                        mCodigoEmail = 0;
                    if (rDocuments["IdEstadoHacienda"] != System.DBNull.Value)
                        mIdEstadoHacienda = (int)rDocuments["IdEstadoHacienda"];
                    else
                        mIdEstadoHacienda = 0;

                    using (StreamWriter writer = new StreamWriter(mClave + ".xml", true))
                    {
                        writer.WriteLine(mXMLConFirma.ToString());
                        writer.Close();                        
                    }

                    if (File.Exists("ConfigPDF.txt"))
                        File.Delete("ConfigPDF.txt");
                    mLinea = Convert.ToString(mIdEmpresa) + ";" + Convert.ToString(mId);
                    using (StreamWriter writer = new StreamWriter("ConfigPDF.txt", true))
                    {
                        writer.WriteLine(mLinea.ToString());
                        writer.Close();
                    }

                    IdPlsntilla = Get_EmpresaIdPlsntilla(mIdEmpresa);
                    Plsntilla = Get_Plantilla(IdPlsntilla);
                    if (File.Exists("Plantilla.fr3"))
                        File.Delete("Plantilla.fr3");
                    using (StreamWriter writer = new StreamWriter("Plantilla.fr3", true))
                    {
                        writer.WriteLine(Plsntilla.ToString());
                        writer.Close();
                    }

                    if (mCodigoEmail == 0)
                    {
                        Process.Start(@"ServicioPDF.exe");
                        Get_EmpresaEmail(dtEmail, mIdEmpresa);
                        SendEmail(dtEmail, mReceptorCorreos, mClave + ".xml", "1.pdf");
                        Actualizar_HaciendaEmailCodigo(1, mId);
                    }
                    if (mIdEstadoHacienda == 0)
                    {
                        string mRespueata = string.Empty;
                        string mDetalle = string.Empty;
                        string mTokenUsuario = string.Empty;
                        string mTokenClave = string.Empty;
                    
                        Get_EmpresaToken(mIdEmpresa, ref mTokenUsuario, ref mTokenClave);
                        Insertar(mClave + ".xml", mClave, ref mRespueata, ref mDetalle, mTokenUsuario, mTokenClave);
                        // Actualizar_Hacienda(1, mId, mRespueata, mDetalle);
                    }
                }
            }
            catch
            { }
        }
        #endregion

        #region Insertar
        private static void Insertar(string pArchivoFirmado, string pClaveDocumento, ref string pRespueataHacienda, ref string pDetalleHacienda, string pTokenUsuario, string pTokenClave)
        {
            try
            {
                string Token = "";
                Token = getToken(pTokenUsuario, pTokenClave);
                    
                XmlDocument xmlElectronica = new XmlDocument();
                xmlElectronica.Load(pArchivoFirmado);

                FacturaElectronicaCR_CS.Recepcion myRecepcion = new FacturaElectronicaCR_CS.Recepcion();

                FacturaElectronicaCR_CS.Receptor myReceptor = new FacturaElectronicaCR_CS.Receptor();
                myReceptor.numeroIdentificacion = "3102007223";
                myReceptor.TipoIdentificacion = "02";
                myReceptor.sinReceptor = false;
                FacturaElectronicaCR_CS.Emisor myEmisor = new FacturaElectronicaCR_CS.Emisor();
                myEmisor.numeroIdentificacion = "3101642839";
                myEmisor.TipoIdentificacion = "02";

                myRecepcion.emisor = myEmisor;
                myRecepcion.receptor = myReceptor;
                myRecepcion.clave = pClaveDocumento;
                myRecepcion.fecha = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz");
                myRecepcion.comprobanteXml = FacturaElectronicaCR_CS.Funciones.EncodeStrToBase64(xmlElectronica.OuterXml);
                xmlElectronica = null;

                FacturaElectronicaCR_CS.Comunicacion enviaFactura = new FacturaElectronicaCR_CS.Comunicacion();
                int mModoProduccion = 1;
                if (mModoProduccion > 0)
                {                    
                    enviaFactura.EnvioDatos(Token, myRecepcion, mModoProduccion);

                    pRespueataHacienda = enviaFactura.pRespueata;
                    pDetalleHacienda = enviaFactura.pDetalle;

                    if (enviaFactura.pRespueata != "Error:400")
                    {
                        string jsonEnvio = "";
                        jsonEnvio = enviaFactura.jsonEnvio;
                        string jsonRespuesta = "";
                        jsonRespuesta = enviaFactura.jsonRespuesta;
                    }
                }
            }
            catch (Exception ex)
            {
               // Escribir_Resouesta("Error:400");
            }
        }
        #endregion

        #region getToken
        public static string getToken(string pTokenUsuario, string pTokenClave)
        {
            try
            {
                FacturaElectronicaCR_CS.TokenHacienda iTokenHacienda = new FacturaElectronicaCR_CS.TokenHacienda();
                iTokenHacienda.GetTokenHacienda(pTokenUsuario, pTokenClave);
                return iTokenHacienda.accessToken;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

    }
}
