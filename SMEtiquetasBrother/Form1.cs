using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OracleClient;
using System.Drawing.Printing;
using MySql.Data.MySqlClient;
using System.IO;
using System.Net.NetworkInformation;
using System.Net;

namespace SMEtiquetasBrother
{
    public partial class FormPrincipal : Form
    {
        public const int QtdEtiquetasPorPagina = 4;

        public string NOME_EMPRESA;
        public string CNPJ;
        public string IE;

        public string Endereco;
        public string Bairro;
        public string Cidade;
        public string UF;
        public string CEP;
        //public string Qtd_Volumes;
        public string OS;

        public string CodigoModelo;
        public string PesoLiquido;
        public string PesoBruto;
        public string CodigoMaterialCliente;
        public string CodigoSM;
        public string NotaFiscal;
        public string DescricaoProduto;
        public string Lote;
        public string NroPedidoCliente;

        public List<string> ListaItem = new List<string>();
        public List<string> ListaNroSerie = new List<string>();

        public MySqlConnection conexao;
        public MySqlCommand comandoSQL;
        public MySqlDataAdapter adaptador;

        public DataSet dataset;

        public bool bTerminouImpressao = false;

        public int PaginaSelecionada = 1;
        public int PaginaAtualImpressao = 0;
        public int QtdItens = 0;
        public int QtdEtiquetas = 0;
        public int QtdEtiquetasColetivas = 4;
        public int QtdPaginasImpressao = 0;

        //For Database connection 
        public OracleConnection conn;

        //To fill DataSet and update Datasource
        private OracleDataAdapter productsAdapter;

        //For automatically generating Commands to make changes to Database through Dataset
        private OracleCommandBuilder productsCmdBuilder;

        //In-memory cache of data
        private DataSet productsDataSet;

        public FormPrincipal()
        {
            InitializeComponent();
        }

        private void FormPrincipal_Load(object sender, EventArgs e)
        {
            String[] BarCodeSplit;
            String BarCodeStr;
            String Qtd_Volumes;

            FormMessage FormMensagem = new FormMessage();

            cbSelecao1Item.SelectedIndex = 0;

            cbTipoCodBarras.SelectedIndex = 5;
            QtdItens = 0;

            tbCliente.Clear();

            this.Text = this.Text + " - " + Application.ProductVersion.ToString();

            // Fazer um ping pra descobrir se é rede interna ou externa.
            clGlobal.MessageWarning = "Aviso";
            clGlobal.MessageText = "Verificando Rede Interna";
            FormMensagem.Show();
            FormMensagem.Refresh();
            System.Threading.Thread.Sleep(5000);
            FormMensagem.Visible = false;
            clGlobal.bRespostaPing = true;
            TestaPing(clGlobal.EnderecoIPPRODSERV);

            while (clGlobal.bTerminouPing == false)
            {
                Application.DoEvents();
            }

            if (clGlobal.bRespostaPing == true)
            {
                if((clGlobal.HostName == clGlobal.EnderecoLocalPRODSERV)||(clGlobal.HostName == clGlobal.EnderecoLocalPRODSERVdominio))
                {
                    clGlobal.EnderecoServidor = clGlobal.EnderecoIPPRODSERV;
                    clGlobal.MessageWarning = "Aviso";
                    clGlobal.MessageText = "Rede Interna - Servidor " + clGlobal.EnderecoLocalPRODSERV + " - IP = " + clGlobal.EnderecoIPPRODSERV;
                    FormMensagem.Visible = true;
                    FormMensagem.Refresh();
                    System.Threading.Thread.Sleep(3000);
                    FormMensagem.Visible = false;
                }
                else
                {
                    clGlobal.EnderecoServidor = clGlobal.EnderecoRemotoPRODSERV;
                    clGlobal.MessageWarning = "Aviso";
                    clGlobal.MessageText = "Rede Externa - Endereço Servidor - " + clGlobal.EnderecoServidor;
                    FormMensagem.Visible = true;
                    FormMensagem.Refresh();
                    System.Threading.Thread.Sleep(3000);
                    FormMensagem.Visible = false;
                }
                tbOrdemProducao.Focus();
            }
            else
            {
                clGlobal.EnderecoServidor = clGlobal.EnderecoRemotoPRODSERV;
                clGlobal.MessageWarning = "Aviso";
                clGlobal.MessageText = "Rede não detectada";
                FormMensagem.Visible = true;
                FormMensagem.Refresh();
                System.Threading.Thread.Sleep(2000);
                FormMensagem.Visible = false;
                this.Close();
            }
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        public Bitmap GerarQRCode(int width, int height, string text)
        {
            try
            {
                var bw = new ZXing.BarcodeWriter();
                var encOptions = new ZXing.Common.EncodingOptions() { Width = width, Height = height, Margin = 0 };
                bw.Options = encOptions;
                bw.Format = ZXing.BarcodeFormat.QR_CODE;
                var resultado = new Bitmap(bw.Write(text));
                return resultado;
            }
            catch
            {
                throw;
            }
        }

        public Bitmap GerarBarCode(int width, int height, string text, string Tipo)
        {
            try
            {
                var bw = new ZXing.BarcodeWriter();
                var encOptions = new ZXing.Common.EncodingOptions() { Width = width, Height = height, Margin = 0 };
                bw.Options = encOptions;
                switch (Tipo)
                {
                    case "EAN-8":
                        bw.Format = ZXing.BarcodeFormat.EAN_8;
                        break;
                    case "EAN-13":
                        bw.Format = ZXing.BarcodeFormat.EAN_13;
                        break;
                    case "UPC-A":
                        bw.Format = ZXing.BarcodeFormat.UPC_A;
                        break;
                    case "UPC-E":
                        bw.Format = ZXing.BarcodeFormat.UPC_E;
                        break;
                    case "CÓDIGO 39":
                        bw.Format = ZXing.BarcodeFormat.CODE_39;
                        break;
                    case "CÓDIGO 128":
                        bw.Format = ZXing.BarcodeFormat.CODE_128;
                        break;
                    case "CODABAR":
                        bw.Format = ZXing.BarcodeFormat.CODABAR;
                        break;
                }
                var resultado = new Bitmap(bw.Write(text));
                return resultado;
            }
            catch
            {
                throw;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String[] BarCodeSplit;
            String BarCodeStr;
            string NomeCampo;
            string NomeTabela;
            string Campo;
            string NroOrdemProducao;
            string NroPedido;

            int IndiceNroSerieInicial = 0;
            
            ListaNroSerie.Clear();
            ListaItem.Clear();

            PaginaSelecionada = 1;
            PaginaAtualImpressao = 0;
            QtdItens = 0;
            QtdEtiquetas = 0;
            QtdPaginasImpressao = 0;

            udPages.Items.Clear();
            pbQRCode1.Image = null;
            pbQRCode1.Visible = false;
            pbQRCode2.Image = null;
            pbQRCode2.Visible = false;
            pbQRCode3.Image = null;
            pbQRCode3.Visible = false;
            pbQRCode4.Image = null;
            pbQRCode4.Visible = false;
            pbQRCode5.Image = null;
            pbQRCode5.Visible = false;
            pbQRCode6.Image = null;
            pbQRCode6.Visible = false;

            pbCodBarras1.Image = null;
            pbCodBarras1.Visible = false;
            pbCodBarras2.Image = null;
            pbCodBarras2.Visible = false;
            pbCodBarras3.Image = null;
            pbCodBarras3.Visible = false;
            pbCodBarras4.Image = null;
            pbCodBarras4.Visible = false;
            pbCodBarras5.Image = null;
            pbCodBarras5.Visible = false;
            pbCodBarras6.Image = null;
            pbCodBarras6.Visible = false;

            try
            {
                NotaFiscal = "";
                Lote = "";

                string connString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" + clGlobal.EnderecoServidor + ")(PORT=1521))(CONNECT_DATA = (SERVER = DEDICATED)(SERVICE_NAME = XE)))";
                //string connString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=mgers.dyndns.org)(PORT=1521))(CONNECT_DATA = (SERVER = DEDICATED)(SERVICE_NAME = XE)))";
                connString = connString + ";User Id=ODIN_CONSULTA;Password=ODIN123@;";
                OracleConnection conn = new OracleConnection(connString);
                conn.Open();
                if(conn.State != ConnectionState.Open)
                {
                    MessageBox.Show("Não foi possível a conexão com o banco ODIN", "Erro");
                    return;
                }

                OracleCommand commandPedidoWord = conn.CreateCommand();

                if ((tbOrdemProducao.Text == "")||(tbOrdemProducao.Text == null))
                {
                    MessageBox.Show("Campo Ordem de Producao inválido");
                    return;
                }

                if ((tbPedidoOdin.Text == "") || (tbPedidoOdin.Text == null))
                {
                    MessageBox.Show("Campo Pedido inválido");
                    return;
                }

                NroOrdemProducao = tbOrdemProducao.Text;
                NroPedido = tbPedidoOdin.Text;

                //string sqlPedido = "SELECT CGC, INSCRICAO_ESTADUAL, ENDERECO, BAIRRO, CEP, UF, QTD_VOLUMES, NRO_OS FROM ODIN_MGE.PEDIDO where NRO_PEDIDO = " + NroOrdemProducao.Text;
                string sqlPedidoWord = "SELECT AGENTE_ENTREGA_NOME, AGENTE_CNPJ, PEDIDO_ENDERECO, PEDIDO_BAIRRO, PEDIDO_CIDADE, PEDIDO_UF, PEDIDO_CEP FROM ODIN_MGE.V_PEDIDO_WORD where NRO_PEDIDO = " + NroPedido;
                //string sqlPedido = "SELECT * FROM ODIN_MGE.V_PEDIDO_WORD WHERE NRO_PEDIDO = " + NroPedido;
                //string sqlPedido = "SELECT * FROM ODIN_MGE.ORDEM_PRODUCAO WHERE NRO_DOCUMENTO = " + NroPedido;
                //string sqlPedido = "SELECT owner, table_name, column_name FROM all_tab_columns where owner = 'ODIN_MGE'";
                //string sqlPedido = "SELECT table_name from all_tables where owner = 'ODIN_MGE'";
                //string sqlPedido = "SELECT owner, table_name, column_name FROM ODIN_MGE";

                commandPedidoWord.CommandText = sqlPedidoWord;

                OracleDataReader readerPedidoWord = commandPedidoWord.ExecuteReader();

                if (!readerPedidoWord.HasRows)
                {
                    MessageBox.Show("Pedido não encontrado");
                    return;
                }

                while (readerPedidoWord.Read())
                {
                    //MessageBox.Show(readerPedido.GetValue(1).ToString() + " - " + readerPedido.GetValue(2).ToString());

                    //for(int i=0; i<readerPedido.FieldCount-1; i++)
                    //{
                    //    MessageBox.Show(readerPedido.GetName(i).ToString() + " - " + readerPedido.GetValue(i).ToString());
                    //}
                    //continue;

                    NOME_EMPRESA = (string)readerPedidoWord.GetString(0).ToString();
                    CNPJ = (string)readerPedidoWord.GetString(1).ToString();
                    Endereco = (string)readerPedidoWord.GetString(2).ToString();
                    Bairro = (string)readerPedidoWord.GetString(3).ToString();
                    Cidade = (string)readerPedidoWord.GetString(4).ToString();
                    UF = (string)readerPedidoWord.GetString(5).ToString();
                    CEP = (string)readerPedidoWord.GetString(6).ToString();

                    //Qtd_Volumes = (string)readerPedido.GetDecimal(6).ToString();
                    //if(readerPedido.GetValue(7) != null)
                    //{
                    //    OS = (string)readerPedido.GetValue(7).ToString();
                    //}
                    //else
                    //{
                    //    OS = "";
                    //}

                    tbCliente.Clear();
                    tbCliente.AppendText("EMPRESA: " + NOME_EMPRESA + "\r\n");
                    tbCliente.AppendText("CNPJ: " + CNPJ + "\r\n");
                    tbCliente.AppendText("Endereço: " + Endereco + "\r\n");
                    tbCliente.AppendText(Bairro + " - " + Cidade + " - " + UF + "\r\n");
                    tbCliente.AppendText("CEP: " + CEP);

                    if ((CNPJ != "")&&(CNPJ != null))
                    {
                        OracleCommand commandPedido = conn.CreateCommand();

                        string sqlPedido = "SELECT * FROM ODIN_MGE.PEDIDO WHERE NRO_DOCUMENTO = " + NroPedido;
                        commandPedido.CommandText = sqlPedido;

                        OracleDataReader readerPedido = commandPedido.ExecuteReader();

                        if (!readerPedido.HasRows)
                        {
                            MessageBox.Show("Pedido não encontrado");
                            return;
                        }

                        while (readerPedido.Read())
                        {
                            if ((readerPedido["NRO_PED_CLI"] != null))
                            {
                                NroPedidoCliente = readerPedido["NRO_PED_CLI"].ToString();
                            }
                            else
                            {
                                NroPedidoCliente = "";
                            }
                        }

                        OracleCommand commandNotaFiscal = conn.CreateCommand();
                        string sqlNotaFiscal = "SELECT * FROM ODIN_MGE.nota where nro_pedido = " + NroPedido;
                        commandNotaFiscal.CommandText = sqlNotaFiscal;

                        OracleDataReader readerNotaFiscal = commandNotaFiscal.ExecuteReader();

                        if (readerNotaFiscal.HasRows)
                        {
                            while (readerNotaFiscal.Read())
                            {
                                if ((readerNotaFiscal["NRO_DOCUMENTO"] != null))
                                {
                                    NotaFiscal = readerNotaFiscal["NRO_DOCUMENTO"].ToString();
                                }
                                else
                                {
                                    NotaFiscal = "";
                                }
                            }
                        }

                        OracleCommand commandCliente = conn.CreateCommand();

                        string sqlCliente = "SELECT * FROM ODIN_MGE.ITEM_PEDIDO WHERE NRO_PEDIDO = " + NroPedido;
                        commandCliente.CommandText = sqlCliente;

                        OracleDataReader readerCliente = commandCliente.ExecuteReader();

                        if (!readerCliente.HasRows)
                        {
                            MessageBox.Show("Pedido não encontrado");
                            return;
                        }

                        while (readerCliente.Read())
                        {
                            //if (((string)readerCliente["COD_ITEM"] != "") && ((string)readerCliente["COD_ITEM"] != null))
                            //{
                            //    ListaItem.Add(readerCliente["COD_ITEM"].ToString());
                            //}

                            //if ((readerCliente["COD_ITEM"] != null))
                            //{
                            //    CodigoSM = readerCliente["COD_ITEM"].ToString();
                            //}
                            //else
                            //{
                            //    CodigoSM = "";
                            //}
                            if ((readerCliente["COD_ITEM_IMPRESSAO"] != null))
                            {
                                CodigoMaterialCliente = readerCliente["COD_ITEM_IMPRESSAO"].ToString();
                            }
                            else
                            {
                                CodigoMaterialCliente = "";
                            }
                            if ((readerCliente["PESO_LIQUIDO"] != null))
                            {
                                PesoLiquido = readerCliente["PESO_LIQUIDO"].ToString();
                            }
                            else
                            {
                                PesoLiquido = "";
                            }

                            if ((readerCliente["PESO_BRUTO"] != null))
                            {
                                PesoBruto = readerCliente["PESO_BRUTO"].ToString();
                            }
                            else
                            {
                                PesoBruto = "";
                            }
                        }
                    }
                }

                tbCodigoSM.Text = CodigoSM;
                tbNotaFiscal.Text = NotaFiscal;
                tbCodMaterial.Text = CodigoMaterialCliente;
                tbPesoLiquido.Text = PesoLiquido;
                tbPesoBruto.Text = PesoBruto;
                tbPedidoSM.Text = NroPedido;
                tbPedido.Text = NroPedidoCliente;
                tbLote.Text = Lote;
                //return;

                // Pegar informação do banco MySQL
                conexao = new MySqlConnection("server= " + clGlobal.EnderecoServidor + "; port=1234; User Id="+clGlobal.UsuarioMySQL+"; database=smp; password="+clGlobal.SenhaMySQL);
                //conexao = new MySqlConnection("server= mgers.dyndns.org; port=1234; User Id=SYSDBA; database=smp; password=masterkey");

                if (conexao.State == ConnectionState.Closed)
                {
                    conexao.Open();
                }

                string TextoAux = "";
                string TextoAux2 = "";
                int TamTexto = 0;
                if (NroOrdemProducao.Length <= 6)
                {
                    TamTexto = 6 - NroOrdemProducao.Length;
                }

                comandoSQL = conexao.CreateCommand();
                TextoAux = NroOrdemProducao;
                for (int z = 0; z < TamTexto; z++)
                {
                    TextoAux2 += "0";
                }

                TextoAux2 += TextoAux;
                NroOrdemProducao = TextoAux2;
                string ComandoAux = "Select * from `smp`.`SM_Produtos` where SM_PD_NroSerie like '" + NroOrdemProducao + "%' ORDER BY SM_PD_NroCliente";
                comandoSQL.CommandText = string.Format(ComandoAux);
                adaptador = new MySqlDataAdapter(comandoSQL);
                dataset = new DataSet();
                adaptador.Fill(dataset);

                if(dataset.Tables[0] == null)
                {
                    MessageBox.Show("Não foi possível abrir o banco SMP.", "Erro");
                    return;
                }

                clGlobal.QtdEtiquetasIndividuais = dataset.Tables[0].Rows.Count;
                if ((clGlobal.QtdEtiquetasIndividuais % QtdEtiquetasColetivas) == 0)
                {
                    clGlobal.QtdEtiquetasColetivas = clGlobal.QtdEtiquetasIndividuais / QtdEtiquetasColetivas;
                }
                else
                {
                    clGlobal.QtdEtiquetasColetivas = (clGlobal.QtdEtiquetasIndividuais / QtdEtiquetasColetivas) + 1;
                }

                if (cbEtiquetaColetiva.Checked)
                {
                    if (cbPedidoParcial.Checked)
                    {
                        if ((tbQtdItensImprimir.Text != null) && (tbQtdItensImprimir.Text != ""))
                        {
                            if((((int.Parse(tbNroInicial.Text) - 1) * int.Parse(tbQtdColetiva.Text)) + (int.Parse(tbQtdItensImprimir.Text)*int.Parse(tbQtdColetiva.Text))) > clGlobal.QtdEtiquetasIndividuais)
                            {
                                QtdItens = clGlobal.QtdEtiquetasIndividuais - ((int.Parse(tbNroInicial.Text) - 1) * int.Parse(tbQtdColetiva.Text));
                            }
                            else
                            {
                                QtdItens = ((int.Parse(tbQtdItensImprimir.Text)) * (int.Parse(tbQtdColetiva.Text)));
                            }
                            //QtdItens = (int.Parse(tbQtdItensImprimir.Text) * int.Parse(tbQtdColetiva.Text));
                        }
                        else
                        {
                            QtdItens = dataset.Tables[0].Rows.Count;
                        }
                    }
                    else
                    {
                        QtdItens = dataset.Tables[0].Rows.Count;
                    }
                }
                else if (cbEtiquetaIndividual.Checked)
                {
                    if (cbPedidoParcial.Checked)
                    {
                        if ((tbQtdItensImprimir.Text != null) && (tbQtdItensImprimir.Text != ""))
                        {
                            QtdItens = int.Parse(tbQtdItensImprimir.Text);
                        }
                        else
                        {
                            QtdItens = dataset.Tables[0].Rows.Count;
                        }
                    }
                    else
                    {
                        QtdItens = dataset.Tables[0].Rows.Count;
                    }
                }

                if (QtdItens == 0)
                {
                    MessageBox.Show("Qtd itens = 0", "Erro");
                    return;
                }

                if (cbEtiquetaIndividual.Checked)
                {
                    QtdEtiquetas = QtdItens;

                    if (QtdEtiquetas > 0)
                    {
                        for (int s = 1; s <= QtdEtiquetas; s++)
                        {
                            udPages.Items.Add(s.ToString());
                        }
                        udPages.SelectedIndex = 0;
                        PaginaSelecionada = int.Parse(udPages.Text);

                        QtdPaginasImpressao = QtdEtiquetas;

                        if (cbPedidoParcial.Checked)
                        {
                            tbCaixa.Text = int.Parse(tbNroInicial.Text).ToString() + "/" + dataset.Tables[0].Rows.Count.ToString();
                        }
                        else
                        {
                            tbCaixa.Text = PaginaSelecionada.ToString() + "/" + QtdEtiquetas;
                        }
                    }
                }
                else
                {
                    if ((QtdItens % QtdEtiquetasColetivas) > 0)
                    {
                        QtdEtiquetas = (QtdItens / QtdEtiquetasColetivas) + 1;
                    }
                    else
                    {
                        QtdEtiquetas = QtdItens / QtdEtiquetasColetivas;
                    }

                    if (QtdEtiquetas > 0)
                    {
                        for (int s = 1; s <= QtdEtiquetas; s++)
                        {
                            udPages.Items.Add(s.ToString());
                        }
                        udPages.SelectedIndex = 0;
                        PaginaSelecionada = int.Parse(udPages.Text);

                        if ((QtdEtiquetas % QtdEtiquetasPorPagina) > 0)
                        {
                            QtdPaginasImpressao = (QtdEtiquetas / QtdEtiquetasPorPagina) + 1;
                        }
                        else
                        {
                            QtdPaginasImpressao = QtdEtiquetas / QtdEtiquetasPorPagina;
                        }

                        //if (cbPedidoParcial.Checked)
                        //{
                        //    tbCaixa.Text = (int.Parse(tbNroInicial.Text) + PaginaSelecionada).ToString() + "/" + QtdEtiquetas;
                        //}
                        //else
                        //{
                            //tbCaixa.Text = PaginaSelecionada.ToString() + "/" + QtdEtiquetas;
                            //tbCaixa.Text = (PaginaSelecionada + (int.Parse(tbNroInicial.Text)/QtdEtiquetasColetivas)).ToString() + "/" + clGlobal.QtdEtiquetasColetivas;
                        //}

                        if (cbPedidoParcial.Checked)
                        {
                            tbCaixa.Text = (PaginaSelecionada + (int.Parse(tbNroInicial.Text)-1)).ToString() + "/" + clGlobal.QtdEtiquetasColetivas;
                        }
                        else
                        {
                            tbCaixa.Text = PaginaSelecionada.ToString() + "/" + clGlobal.QtdEtiquetasColetivas;
                        }
                    }
                }

                if (QtdItens > 0)
                {
                    if (cbEtiquetaColetiva.Checked)
                    {

                        if (cbPedidoParcial.Checked)
                        {
                            if ((tbNroInicial.Text != null) && (tbNroInicial.Text != ""))
                            {
                                IndiceNroSerieInicial = ((((int.Parse(tbNroInicial.Text) - 1) * int.Parse(tbQtdColetiva.Text)) + 1) - 1);
                            }
                            else
                            {
                                IndiceNroSerieInicial = 0;
                            }
                        }
                        else
                        {
                            IndiceNroSerieInicial = 0;
                        }
                    }
                    else if(cbEtiquetaIndividual.Checked)
                    {
                        if (cbPedidoParcial.Checked)
                        {
                            if ((tbNroInicial.Text != null) && (tbNroInicial.Text != ""))
                            {
                                IndiceNroSerieInicial = (int.Parse(tbNroInicial.Text) - 1);
                            }
                            else
                            {
                                IndiceNroSerieInicial = 0;
                            }
                        }
                        else
                        {
                            IndiceNroSerieInicial = 0;
                        }
                    }

                    for (int x = IndiceNroSerieInicial; x < (IndiceNroSerieInicial + QtdItens); x++)
                    {
                        ListaNroSerie.Add(dataset.Tables[0].Rows[x]["SM_PD_StrCliente"].ToString());
                    }
                }

                adaptador.Dispose();
                dataset.Dispose();
                conexao.Dispose();
                comandoSQL.Dispose();

                // Pegar dados do modelo
                CodigoModelo = "";
                ComandoAux = "Select * from `smp`.`SM_OrdemProducao` where SM_OP_Numero like '" + NroOrdemProducao + "%'";
                comandoSQL.CommandText = string.Format(ComandoAux);
                adaptador = new MySqlDataAdapter(comandoSQL);
                dataset = new DataSet();
                adaptador.Fill(dataset);

                if(dataset.Tables[0].Rows.Count > 0)
                {
                    CodigoModelo = dataset.Tables[0].Rows[0]["SM_OP_CodModelo"].ToString();
                }

                adaptador.Dispose();
                dataset.Dispose();
                conexao.Dispose();
                comandoSQL.Dispose();

                if(CodigoModelo != "")
                {
                    ComandoAux = "Select * from `smp`.`SM_Modelos` where SM_MD_Codigo = '" + CodigoModelo + "'";
                    comandoSQL.CommandText = string.Format(ComandoAux);
                    adaptador = new MySqlDataAdapter(comandoSQL);
                    dataset = new DataSet();
                    adaptador.Fill(dataset);

                    if (dataset.Tables[0] == null)
                    {
                        MessageBox.Show("Não foi possível abrir banco de modelos de equipamentos", "Erro");
                        return;
                    }

                    if (dataset.Tables[0].Rows.Count > 0)
                    {
                        tbDescricaoProduto.Text = /*QtdItens.ToString() + " - " + */dataset.Tables[0].Rows[0]["SM_MD_Descricao"].ToString();
                        tbCodigoSM.Text = dataset.Tables[0].Rows[0]["SM_MD_Codigo"].ToString();
                    }
                    else
                    {
                        tbDescricaoProduto.Text = "";
                        tbCodigoSM.Text = "";
                    }

                    adaptador.Dispose();
                    dataset.Dispose();
                    conexao.Dispose();
                    comandoSQL.Dispose();
                }

                // Colocar na tela os numeros de série
                if (ListaNroSerie != null)
                {
                    if (cbEtiquetaColetiva.Checked)
                    {
                        if (QtdEtiquetasColetivas == 2)
                        {
                            if (ListaNroSerie[0] != null)
                            {
                                BarCodeSplit = ListaNroSerie[0].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[0];
                                }

                                pbQRCode1.Image = GerarQRCode(pbQRCode1.Width, pbQRCode1.Height, BarCodeStr);
                                pbCodBarras1.Image = GerarBarCode(pbCodBarras1.Width, pbCodBarras1.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode1.Visible = true;
                                pbCodBarras1.Visible = true;
                            }
                            else
                            {
                                pbQRCode1.Visible = false;
                                pbCodBarras1.Visible = false;
                            }

                            if (ListaNroSerie[1] != null)
                            {
                                BarCodeSplit = ListaNroSerie[1].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[1];
                                }

                                pbQRCode2.Image = GerarQRCode(pbQRCode2.Width, pbQRCode2.Height, BarCodeStr);
                                pbCodBarras2.Image = GerarBarCode(pbCodBarras2.Width, pbCodBarras2.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode2.Visible = true;
                                pbCodBarras2.Visible = true;
                            }
                            else
                            {
                                pbQRCode2.Visible = false;
                                pbCodBarras2.Visible = false;
                            }
                        }
                        else if (QtdEtiquetasColetivas == 3)
                        {
                            if (ListaNroSerie[0] != null)
                            {
                                BarCodeSplit = ListaNroSerie[0].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[0];
                                }

                                pbQRCode1.Image = GerarQRCode(pbQRCode1.Width, pbQRCode1.Height, BarCodeStr);
                                pbCodBarras1.Image = GerarBarCode(pbCodBarras1.Width, pbCodBarras1.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode1.Visible = true;
                                pbCodBarras1.Visible = true;
                            }
                            else
                            {
                                pbQRCode1.Visible = false;
                                pbCodBarras1.Visible = false;
                            }

                            if (ListaNroSerie[1] != null)
                            {
                                BarCodeSplit = ListaNroSerie[1].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[1];
                                }

                                pbQRCode2.Image = GerarQRCode(pbQRCode2.Width, pbQRCode2.Height, BarCodeStr);
                                pbCodBarras2.Image = GerarBarCode(pbCodBarras2.Width, pbCodBarras2.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode2.Visible = true;
                                pbCodBarras2.Visible = true;
                            }
                            else
                            {
                                pbQRCode2.Visible = false;
                                pbCodBarras2.Visible = false;
                            }

                            if (ListaNroSerie[2] != null)
                            {
                                BarCodeSplit = ListaNroSerie[2].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[2];
                                }

                                pbQRCode3.Image = GerarQRCode(pbQRCode3.Width, pbQRCode3.Height, BarCodeStr);
                                pbCodBarras3.Image = GerarBarCode(pbCodBarras3.Width, pbCodBarras3.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode3.Visible = true;
                                pbCodBarras3.Visible = true;
                            }
                            else
                            {
                                pbQRCode3.Visible = false;
                                pbCodBarras3.Visible = false;
                            }
                        }
                        else if (QtdEtiquetasColetivas == 4)
                        {
                            if (ListaNroSerie[0] != null)
                            {
                                BarCodeSplit = ListaNroSerie[0].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[0];
                                }

                                pbQRCode1.Image = GerarQRCode(pbQRCode1.Width, pbQRCode1.Height, BarCodeStr);
                                pbCodBarras1.Image = GerarBarCode(pbCodBarras1.Width, pbCodBarras1.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode1.Visible = true;
                                pbCodBarras1.Visible = true;
                            }
                            else
                            {
                                pbQRCode1.Visible = false;
                                pbCodBarras1.Visible = false;
                            }

                            if (ListaNroSerie[1] != null)
                            {
                                BarCodeSplit = ListaNroSerie[1].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[1];
                                }

                                pbQRCode2.Image = GerarQRCode(pbQRCode2.Width, pbQRCode2.Height, BarCodeStr);
                                pbCodBarras2.Image = GerarBarCode(pbCodBarras2.Width, pbCodBarras2.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode2.Visible = true;
                                pbCodBarras2.Visible = true;
                            }
                            else
                            {
                                pbQRCode2.Visible = false;
                                pbCodBarras2.Visible = false;
                            }

                            if (ListaNroSerie[2] != null)
                            {
                                BarCodeSplit = ListaNroSerie[2].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[2];
                                }

                                pbQRCode3.Image = GerarQRCode(pbQRCode3.Width, pbQRCode3.Height, BarCodeStr);
                                pbCodBarras3.Image = GerarBarCode(pbCodBarras3.Width, pbCodBarras3.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode3.Visible = true;
                                pbCodBarras3.Visible = true;
                            }
                            else
                            {
                                pbQRCode3.Visible = false;
                                pbCodBarras3.Visible = false;
                            }

                            if (ListaNroSerie[3] != null)
                            {
                                BarCodeSplit = ListaNroSerie[3].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[3];
                                }

                                pbQRCode4.Image = GerarQRCode(pbQRCode4.Width, pbQRCode4.Height, BarCodeStr);
                                pbCodBarras4.Image = GerarBarCode(pbCodBarras4.Width, pbCodBarras4.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode4.Visible = true;
                                pbCodBarras4.Visible = true;
                            }
                            else
                            {
                                pbQRCode4.Visible = false;
                                pbCodBarras4.Visible = false;
                            }
                        }
                        else if (QtdEtiquetasColetivas == 5)
                        {
                            if (ListaNroSerie[0] != null)
                            {
                                BarCodeSplit = ListaNroSerie[0].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[0];
                                }

                                pbQRCode1.Image = GerarQRCode(pbQRCode1.Width, pbQRCode1.Height, BarCodeStr);
                                pbCodBarras1.Image = GerarBarCode(pbCodBarras1.Width, pbCodBarras1.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode1.Visible = true;
                                pbCodBarras1.Visible = true;
                            }
                            else
                            {
                                pbQRCode1.Visible = false;
                                pbCodBarras1.Visible = false;
                            }

                            if (ListaNroSerie[1] != null)
                            {
                                BarCodeSplit = ListaNroSerie[1].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[1];
                                }

                                pbQRCode2.Image = GerarQRCode(pbQRCode2.Width, pbQRCode2.Height, BarCodeStr);
                                pbCodBarras2.Image = GerarBarCode(pbCodBarras2.Width, pbCodBarras2.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode2.Visible = true;
                                pbCodBarras2.Visible = true;
                            }
                            else
                            {
                                pbQRCode2.Visible = false;
                                pbCodBarras2.Visible = false;
                            }

                            if (ListaNroSerie[2] != null)
                            {
                                BarCodeSplit = ListaNroSerie[2].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[2];
                                }

                                pbQRCode3.Image = GerarQRCode(pbQRCode3.Width, pbQRCode3.Height, BarCodeStr);
                                pbCodBarras3.Image = GerarBarCode(pbCodBarras3.Width, pbCodBarras3.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode3.Visible = true;
                                pbCodBarras3.Visible = true;
                            }
                            else
                            {
                                pbQRCode3.Visible = false;
                                pbCodBarras3.Visible = false;
                            }

                            if (ListaNroSerie[3] != null)
                            {
                                BarCodeSplit = ListaNroSerie[3].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[3];
                                }

                                pbQRCode4.Image = GerarQRCode(pbQRCode4.Width, pbQRCode4.Height, BarCodeStr);
                                pbCodBarras4.Image = GerarBarCode(pbCodBarras4.Width, pbCodBarras4.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode4.Visible = true;
                                pbCodBarras4.Visible = true;
                            }
                            else
                            {
                                pbQRCode4.Visible = false;
                                pbCodBarras4.Visible = false;
                            }

                            if (ListaNroSerie[4] != null)
                            {
                                BarCodeSplit = ListaNroSerie[4].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[4];
                                }

                                pbQRCode5.Image = GerarQRCode(pbQRCode5.Width, pbQRCode5.Height, BarCodeStr);
                                pbCodBarras5.Image = GerarBarCode(pbCodBarras5.Width, pbCodBarras5.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode5.Visible = true;
                                pbCodBarras5.Visible = true;
                            }
                            else
                            {
                                pbQRCode5.Visible = false;
                                pbCodBarras5.Visible = false;
                            }
                        }
                        else if(QtdEtiquetasColetivas == 6)
                        {
                            if (ListaNroSerie[0] != null)
                            {
                                BarCodeSplit = ListaNroSerie[0].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[0];
                                }

                                pbQRCode1.Image = GerarQRCode(pbQRCode1.Width, pbQRCode1.Height, BarCodeStr);
                                pbCodBarras1.Image = GerarBarCode(pbCodBarras1.Width, pbCodBarras1.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode1.Visible = true;
                                pbCodBarras1.Visible = true;
                            }
                            else
                            {
                                pbQRCode1.Visible = false;
                                pbCodBarras1.Visible = false;
                            }

                            if (ListaNroSerie[1] != null)
                            {
                                BarCodeSplit = ListaNroSerie[1].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[1];
                                }

                                pbQRCode2.Image = GerarQRCode(pbQRCode2.Width, pbQRCode2.Height, BarCodeStr);
                                pbCodBarras2.Image = GerarBarCode(pbCodBarras2.Width, pbCodBarras2.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode2.Visible = true;
                                pbCodBarras2.Visible = true;
                            }
                            else
                            {
                                pbQRCode2.Visible = false;
                                pbCodBarras2.Visible = false;
                            }

                            if (ListaNroSerie[2] != null)
                            {
                                BarCodeSplit = ListaNroSerie[2].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[2];
                                }

                                pbQRCode3.Image = GerarQRCode(pbQRCode3.Width, pbQRCode3.Height, BarCodeStr);
                                pbCodBarras3.Image = GerarBarCode(pbCodBarras3.Width, pbCodBarras3.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode3.Visible = true;
                                pbCodBarras3.Visible = true;
                            }
                            else
                            {
                                pbQRCode3.Visible = false;
                                pbCodBarras3.Visible = false;
                            }

                            if (ListaNroSerie[3] != null)
                            {
                                BarCodeSplit = ListaNroSerie[3].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[3];
                                }

                                pbQRCode4.Image = GerarQRCode(pbQRCode4.Width, pbQRCode4.Height, BarCodeStr);
                                pbCodBarras4.Image = GerarBarCode(pbCodBarras4.Width, pbCodBarras4.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode4.Visible = true;
                                pbCodBarras4.Visible = true;
                            }
                            else
                            {
                                pbQRCode4.Visible = false;
                                pbCodBarras4.Visible = false;
                            }

                            if (ListaNroSerie[4] != null)
                            {
                                BarCodeSplit = ListaNroSerie[4].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[4];
                                }

                                pbQRCode5.Image = GerarQRCode(pbQRCode5.Width, pbQRCode5.Height, BarCodeStr);
                                pbCodBarras5.Image = GerarBarCode(pbCodBarras5.Width, pbCodBarras5.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode5.Visible = true;
                                pbCodBarras5.Visible = true;
                            }
                            else
                            {
                                pbQRCode5.Visible = false;
                                pbCodBarras5.Visible = false;
                            }

                            if (ListaNroSerie[5] != null)
                            {
                                BarCodeSplit = ListaNroSerie[5].Split('-');
                                if (BarCodeSplit.Length > 1)
                                {
                                    BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                                }
                                else
                                {
                                    BarCodeStr = ListaNroSerie[5];
                                }

                                pbQRCode6.Image = GerarQRCode(pbQRCode6.Width, pbQRCode6.Height, BarCodeStr);
                                pbCodBarras6.Image = GerarBarCode(pbCodBarras6.Width, pbCodBarras6.Height, BarCodeStr, cbTipoCodBarras.Text);
                                //pbQRCode6.Visible = true;
                                pbCodBarras6.Visible = true;
                            }
                            else
                            {
                                pbQRCode6.Visible = false;
                                pbCodBarras6.Visible = false;
                            }
                        }
                    }
                    else
                    {
                        if (ListaNroSerie[0] != null)
                        {
                            BarCodeSplit = ListaNroSerie[0].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[0];
                            }

                            pbQRCode1.Image = GerarQRCode(pbQRCode1.Width, pbQRCode1.Height, BarCodeStr);
                            pbCodBarras1.Image = GerarBarCode(pbCodBarras1.Width, pbCodBarras1.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode1.Visible = true;
                            pbCodBarras1.Visible = true;
                        }
                        else
                        {
                            pbQRCode1.Visible = false;
                            pbCodBarras1.Visible = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //MessageBox.Show("Qtd Etiquetas: " + QtdEtiquetas.ToString(), "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tbItem_TextChanged(object sender, EventArgs e)
        {

        }

        private void imprimirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Call ShowDialog  

            bTerminouImpressao = false;

            printDocument1.DocumentName = "Etiqueta MGE";
            Font f = new Font("Courier New", 10);
            ////  Variável para armazenamento de posicao vertical.
            int posY = printDocument1.DefaultPageSettings.Margins.Top;
            printDocument1.DefaultPageSettings.Landscape = false;
            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                printDocument1.Print();
            }
        }

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            float QtdCodigoBarrasEtiqueta = 0;

            float posicaoX;
            float posicaoY;
            float Intervalo;

            float posicaoXPedidoSM;
            float posicaoYPedidoSM;

            float posicaoXCodigoSM;
            float posicaoYCodigoSM;

            float posicaoXCodigoMaterial;
            float posicaoYCodigoMaterial;

            float posicaoXPedido;
            float posicaoYPedido;

            float posicaoXNotaFiscal;
            float posicaoYNotaFiscal;

            float posicaoXDescricao;
            float posicaoYDescricao;

            float posicaoXItem;
            float posicaoYItem;

            float posicaoXLote;
            float posicaoYLote;

            float posicaoXCaixa;
            float posicaoYCaixa;

            float posicaoXPesoLiquido;
            float posicaoYPesoLiquido;

            float posicaoXPesoBruto;
            float posicaoYPesoBruto;

            TextReader readtbCliente = new System.IO.StringReader(tbCliente.Text);
            int rows = 10;

            string[] text1 = new string[rows];
            for (int r = 0; r < rows-1; r++)
            {
                text1[r] = readtbCliente.ReadLine();
            }

            string ClienteLinha1 = text1[0];
            string ClienteLinha2 = text1[1];
            string ClienteLinha3 = text1[2];
            string ClienteLinha4 = text1[3];
            string ClienteLinha5 = text1[4];
            string ClienteLinha6 = text1[5];

            string DescricaoLinha1 = "";
            string DescricaoLinha2 = "";
            string DescricaoLinha3 = "";
            string DescricaoLinha4 = "";

            if (tbDescricaoProduto.Text.Length > 30)
            {
                DescricaoLinha1 = tbDescricaoProduto.Text.Substring(0, 30);
                if (tbDescricaoProduto.Text.Length > 60)
                {
                    DescricaoLinha2 = tbDescricaoProduto.Text.Substring(30, 30);
                    if (tbDescricaoProduto.Text.Length > 90)
                    {
                        DescricaoLinha3 = tbDescricaoProduto.Text.Substring(60, 30);
                        if(tbDescricaoProduto.Text.Length > 120)
                        {
                            DescricaoLinha4 = tbDescricaoProduto.Text.Substring(90, 30);
                        }
                        else
                        {
                            DescricaoLinha4 = tbDescricaoProduto.Text.Substring(90, tbDescricaoProduto.Text.Length-90);
                        }
                    }
                    else
                    {
                        DescricaoLinha3 = tbDescricaoProduto.Text.Substring(60, tbDescricaoProduto.Text.Length-60);
                    }
                }
                else
                {
                    DescricaoLinha2 = tbDescricaoProduto.Text.Substring(30, tbDescricaoProduto.Text.Length-30);
                }
            }
            else
            {
                DescricaoLinha1 = tbDescricaoProduto.Text;
            }

            Font fonte = new Font(tbCliente.Font.Name, 8);
            SizeF tamanho = e.Graphics.MeasureString(ClienteLinha1, fonte);

            Font fonteEmpresa = new Font(tbCliente.Font.Name, 8);
            SizeF tamanhoEmpresa = e.Graphics.MeasureString(ClienteLinha1, fonteEmpresa);

            posicaoX = e.PageBounds.Left + 16;
            posicaoY = e.PageBounds.Top + 84;

            posicaoXPedidoSM = posicaoX;
            posicaoYPedidoSM = posicaoY + 84;

            posicaoXCodigoSM = posicaoX;
            posicaoYCodigoSM = posicaoY + 123;

            posicaoXCodigoMaterial = posicaoX;
            posicaoYCodigoMaterial = posicaoY + 163;

            posicaoXPedido = posicaoX + 105;
            posicaoYPedido = posicaoY + 84;

            posicaoXNotaFiscal = posicaoX + 105;
            posicaoYNotaFiscal = posicaoY + 123;

            posicaoXDescricao = posicaoX + 218;
            posicaoYDescricao = posicaoY + 84;

            posicaoXItem = posicaoX + 423;
            posicaoYItem = posicaoY + 5;

            posicaoXLote = posicaoX + 422;
            posicaoYLote = posicaoY + 44;

            posicaoXCaixa = posicaoX + 422;
            posicaoYCaixa = posicaoY + 84;

            posicaoXPesoLiquido = posicaoX + 422;
            posicaoYPesoLiquido = posicaoY + 123;

            posicaoXPesoBruto = posicaoX + 422;
            posicaoYPesoBruto = posicaoY + 161;

            Intervalo = 0;

            if(cbSelecao1Item.Text == "TODOS")
            {
                if (cbEtiquetaColetiva.Checked)
                {
                    for (int s = (QtdEtiquetasPorPagina * PaginaAtualImpressao) + 1; s <= QtdEtiquetas; s++)
                    {
                        QtdCodigoBarrasEtiqueta = 0;
                        if (pbCodBarras1.Visible)
                        {
                            QtdCodigoBarrasEtiqueta++;
                        }
                        if (pbCodBarras2.Visible)
                        {
                            QtdCodigoBarrasEtiqueta++;
                        }
                        if (pbCodBarras3.Visible)
                        {
                            QtdCodigoBarrasEtiqueta++;
                        }
                        if (pbCodBarras4.Visible)
                        {
                            QtdCodigoBarrasEtiqueta++;
                        }
                        if (pbCodBarras5.Visible)
                        {
                            QtdCodigoBarrasEtiqueta++;
                        }
                        if (pbCodBarras6.Visible)
                        {
                            QtdCodigoBarrasEtiqueta++;
                        }
                        if (ClienteLinha1 != null)
                        {
                            e.Graphics.DrawString(ClienteLinha1, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo);
                        }

                        if (ClienteLinha2 != null)
                        {
                            e.Graphics.DrawString(ClienteLinha2, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 12);
                        }

                        if (ClienteLinha3 != null)
                        {
                            e.Graphics.DrawString(ClienteLinha3, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 24);
                        }

                        if (ClienteLinha4 != null)
                        {
                            e.Graphics.DrawString(ClienteLinha4, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 36);
                        }

                        if (ClienteLinha5 != null)
                        {
                            e.Graphics.DrawString(ClienteLinha5, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 48);
                        }

                        if (ClienteLinha6 != null)
                        {
                            e.Graphics.DrawString(ClienteLinha6, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 60);
                        }

                        if (tbPedidoSM.Text != "")
                        {
                            e.Graphics.DrawString(tbPedidoSM.Text, fonte, new SolidBrush(Color.Black), posicaoXPedidoSM, posicaoYPedidoSM + Intervalo);
                        }

                        if (tbCodigoSM.Text != "")
                        {
                            e.Graphics.DrawString(tbCodigoSM.Text, fonte, new SolidBrush(Color.Black), posicaoXCodigoSM, posicaoYCodigoSM + Intervalo);
                        }

                        if (tbCodMaterial.Text != "")
                        {
                            e.Graphics.DrawString(tbCodMaterial.Text, fonte, new SolidBrush(Color.Black), posicaoXCodigoMaterial, posicaoYCodigoMaterial + Intervalo);
                        }

                        if (tbPedido.Text != "")
                        {
                            e.Graphics.DrawString(tbPedido.Text, fonte, new SolidBrush(Color.Black), posicaoXPedido, posicaoYPedido + Intervalo);
                        }

                        if (DescricaoLinha1 != null)
                        {
                            e.Graphics.DrawString(DescricaoLinha1, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo);
                        }

                        if (DescricaoLinha2 != null)
                        {
                            e.Graphics.DrawString(DescricaoLinha2, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo + 12);
                        }

                        if (DescricaoLinha3 != null)
                        {
                            e.Graphics.DrawString(DescricaoLinha3, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo + 24);
                        }

                        if (DescricaoLinha4 != null)
                        {
                            e.Graphics.DrawString(DescricaoLinha4, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo + 36);
                        }

                        //if (tbDescricaoProduto.Text != "")
                        //{
                        //    e.Graphics.DrawString(tbDescricaoProduto.Text, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo);
                        //}

                        if (tbNotaFiscal.Text != "")
                        {
                            e.Graphics.DrawString(tbNotaFiscal.Text, fonte, new SolidBrush(Color.Black), posicaoXNotaFiscal, posicaoYNotaFiscal + Intervalo);
                        }

                        if (tbItem.Text != "")
                        {
                            e.Graphics.DrawString(tbItem.Text, fonte, new SolidBrush(Color.Black), posicaoXItem, posicaoYItem + Intervalo);
                        }

                        if (tbLote.Text != "")
                        {
                            e.Graphics.DrawString(tbLote.Text, fonte, new SolidBrush(Color.Black), posicaoXLote, posicaoYLote + Intervalo);
                        }

                        if (tbCaixa.Text != "")
                        {
                            e.Graphics.DrawString(tbCaixa.Text, fonte, new SolidBrush(Color.Black), posicaoXCaixa, posicaoYCaixa + Intervalo);
                        }

                        if (tbPesoLiquido.Text != "")
                        {
                            e.Graphics.DrawString((float.Parse(tbPesoLiquido.Text)*QtdCodigoBarrasEtiqueta).ToString("F" + 4), fonte, new SolidBrush(Color.Black), posicaoXPesoLiquido, posicaoYPesoLiquido + Intervalo);
                        }

                        if (tbPesoBruto.Text != "")
                        {
                            e.Graphics.DrawString((float.Parse(tbPesoBruto.Text)*QtdCodigoBarrasEtiqueta).ToString("F" + 4), fonte, new SolidBrush(Color.Black), posicaoXPesoBruto, posicaoYPesoBruto + Intervalo);
                        }

                        if (QtdEtiquetasColetivas == 2)
                        {
                            if (pbQRCode1.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode1.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 90, pbQRCode1.Image.Width, pbQRCode1.Image.Height);
                            }
                            if (pbCodBarras1.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras1.Image, (e.PageBounds.Right - 330 + pbQRCode1.Image.Width + 10), e.PageBounds.Top + Intervalo + 90, pbCodBarras1.Image.Width, pbCodBarras1.Image.Height);
                            }

                            if (pbQRCode2.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode2.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 135, pbQRCode2.Image.Width, pbQRCode2.Image.Height);
                            }
                            if (pbCodBarras2.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras2.Image, (e.PageBounds.Right - 330 + pbQRCode2.Image.Width + 10), e.PageBounds.Top + Intervalo + 135, pbCodBarras2.Image.Width, pbCodBarras2.Image.Height);
                            }
                        }
                        else if (QtdEtiquetasColetivas == 3)
                        {
                            if (pbQRCode1.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode1.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 90, pbQRCode1.Image.Width, pbQRCode1.Image.Height);
                            }
                            if (pbCodBarras1.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras1.Image, (e.PageBounds.Right - 330 + pbQRCode1.Image.Width + 10), e.PageBounds.Top + Intervalo + 90, pbCodBarras1.Image.Width, pbCodBarras1.Image.Height);
                            }

                            if (pbQRCode2.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode2.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 135, pbQRCode2.Image.Width, pbQRCode2.Image.Height);
                            }
                            if (pbCodBarras2.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras2.Image, (e.PageBounds.Right - 330 + pbQRCode2.Image.Width + 10), e.PageBounds.Top + Intervalo + 135, pbCodBarras2.Image.Width, pbCodBarras2.Image.Height);
                            }
                            if (pbQRCode3.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode3.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 180, pbQRCode3.Image.Width, pbQRCode3.Image.Height);
                            }
                            if (pbCodBarras3.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras3.Image, (e.PageBounds.Right - 330 + pbQRCode3.Image.Width + 10), e.PageBounds.Top + Intervalo + 180, pbCodBarras3.Image.Width, pbCodBarras3.Image.Height);
                            }
                        }
                        else if (QtdEtiquetasColetivas == 4)
                        {
                            if (pbQRCode1.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode1.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 90, pbQRCode1.Image.Width, pbQRCode1.Image.Height);
                            }
                            if (pbCodBarras1.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras1.Image, (e.PageBounds.Right - 330 + pbQRCode1.Image.Width + 10), e.PageBounds.Top + Intervalo + 90, pbCodBarras1.Image.Width, pbCodBarras1.Image.Height);
                            }

                            if (pbQRCode2.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode2.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 135, pbQRCode2.Image.Width, pbQRCode2.Image.Height);
                            }
                            if (pbCodBarras2.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras2.Image, (e.PageBounds.Right - 330 + pbQRCode2.Image.Width + 10), e.PageBounds.Top + Intervalo + 135, pbCodBarras2.Image.Width, pbCodBarras2.Image.Height);
                            }
                            if (pbQRCode3.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode3.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 180, pbQRCode3.Image.Width, pbQRCode3.Image.Height);
                            }
                            if (pbCodBarras3.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras3.Image, (e.PageBounds.Right - 330 + pbQRCode3.Image.Width + 10), e.PageBounds.Top + Intervalo + 180, pbCodBarras3.Image.Width, pbCodBarras3.Image.Height);
                            }
                            if (pbQRCode4.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode4.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 225, pbQRCode4.Image.Width, pbQRCode4.Image.Height);
                            }
                            if (pbCodBarras4.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras4.Image, (e.PageBounds.Right - 330 + pbQRCode4.Image.Width + 10), e.PageBounds.Top + Intervalo + 225, pbCodBarras4.Image.Width, pbCodBarras4.Image.Height);
                            }
                        }
                        else if (QtdEtiquetasColetivas == 5)
                        {
                            if (pbQRCode1.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode1.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 90, pbQRCode1.Image.Width, pbQRCode1.Image.Height * 2 / 3);
                            }
                            if (pbCodBarras1.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras1.Image, (e.PageBounds.Right - 330 + pbQRCode1.Image.Width + 10), e.PageBounds.Top + Intervalo + 90, pbCodBarras1.Image.Width, pbCodBarras1.Image.Height * 2 / 3);
                            }
                            if (pbQRCode2.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode2.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 117, pbQRCode2.Image.Width, pbQRCode2.Image.Height * 2 / 3);
                            }
                            if (pbCodBarras2.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras2.Image, (e.PageBounds.Right - 330 + pbQRCode2.Image.Width + 10), e.PageBounds.Top + Intervalo + 117, pbCodBarras2.Image.Width, pbCodBarras2.Image.Height * 2 / 3);
                            }
                            if (pbQRCode3.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode3.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 144, pbQRCode3.Image.Width, pbQRCode3.Image.Height * 2 / 3);
                            }
                            if (pbCodBarras3.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras3.Image, (e.PageBounds.Right - 330 + pbQRCode3.Image.Width + 10), e.PageBounds.Top + Intervalo + 144, pbCodBarras3.Image.Width, pbCodBarras3.Image.Height * 2 / 3);
                            }
                            if (pbQRCode4.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode4.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 171, pbQRCode4.Image.Width, pbQRCode4.Image.Height * 2 / 3);
                            }
                            if (pbCodBarras4.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras4.Image, (e.PageBounds.Right - 330 + pbQRCode4.Image.Width + 10), e.PageBounds.Top + Intervalo + 171, pbCodBarras4.Image.Width, pbCodBarras4.Image.Height * 2 / 3);
                            }
                            if (pbQRCode5.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode5.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 198, pbQRCode5.Image.Width, pbQRCode5.Image.Height * 2 / 3);
                            }
                            if (pbCodBarras5.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras5.Image, (e.PageBounds.Right - 330 + pbQRCode5.Image.Width + 10), e.PageBounds.Top + Intervalo + 198, pbCodBarras5.Image.Width, pbCodBarras5.Image.Height * 2 / 3);
                            }
                        }
                        else if (QtdEtiquetasColetivas == 6)
                        {
                            if (pbQRCode1.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode1.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 90, pbQRCode1.Image.Width, pbQRCode1.Image.Height*2/3);
                            }
                            if (pbCodBarras1.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras1.Image, (e.PageBounds.Right - 330 + pbQRCode1.Image.Width + 10), e.PageBounds.Top + Intervalo + 90, pbCodBarras1.Image.Width, pbCodBarras1.Image.Height*2/3);
                            }
                            if (pbQRCode2.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode2.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 117, pbQRCode2.Image.Width, pbQRCode2.Image.Height*2/3);
                            }
                            if (pbCodBarras2.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras2.Image, (e.PageBounds.Right - 330 + pbQRCode2.Image.Width + 10), e.PageBounds.Top + Intervalo + 117, pbCodBarras2.Image.Width, pbCodBarras2.Image.Height*2/3);
                            }
                            if (pbQRCode3.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode3.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 144, pbQRCode3.Image.Width, pbQRCode3.Image.Height*2/3);
                            }
                            if (pbCodBarras3.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras3.Image, (e.PageBounds.Right - 330 + pbQRCode3.Image.Width + 10), e.PageBounds.Top + Intervalo + 144, pbCodBarras3.Image.Width, pbCodBarras3.Image.Height*2/3);
                            }
                            if (pbQRCode4.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode4.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 171, pbQRCode4.Image.Width, pbQRCode4.Image.Height*2/3);
                            }
                            if (pbCodBarras4.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras4.Image, (e.PageBounds.Right - 330 + pbQRCode4.Image.Width + 10), e.PageBounds.Top + Intervalo + 171, pbCodBarras4.Image.Width, pbCodBarras4.Image.Height*2/3);
                            }
                            if (pbQRCode5.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode5.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 198, pbQRCode5.Image.Width, pbQRCode5.Image.Height*2/3);
                            }
                            if (pbCodBarras5.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras5.Image, (e.PageBounds.Right - 330 + pbQRCode5.Image.Width + 10), e.PageBounds.Top + Intervalo + 198, pbCodBarras5.Image.Width, pbCodBarras5.Image.Height*2/3);
                            }
                            if (pbQRCode6.Visible)
                            {
                                e.Graphics.DrawImage(pbQRCode6.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 225, pbQRCode6.Image.Width, pbQRCode6.Image.Height*2/3);
                            }
                            if (pbCodBarras6.Visible)
                            {
                                e.Graphics.DrawImage(pbCodBarras6.Image, (e.PageBounds.Right - 330 + pbQRCode6.Image.Width + 10), e.PageBounds.Top + Intervalo + 225, pbCodBarras6.Image.Width, pbCodBarras6.Image.Height*2/3);
                            }
                        }
                        udPages.DownButton();
                        Intervalo += 274;

                        if (s == QtdEtiquetas)
                        {
                            bTerminouImpressao = true;
                        }

                        if (((s % QtdEtiquetasPorPagina) == 0)&&(!bTerminouImpressao))
                        {
                            // Já imprimou 4 etiquetas
                            // Tem que trocar de página
                            PaginaAtualImpressao++;
                            e.HasMorePages = true;
                            break;
                        }
                    }
                }
                else
                {
                    // Etiqueta Individual
                    for (int s = (QtdEtiquetasPorPagina * PaginaAtualImpressao) + 1; s <= QtdEtiquetas; s++)
                    {
                        QtdCodigoBarrasEtiqueta = 0;
                        if (pbCodBarras1.Visible)
                        {
                            QtdCodigoBarrasEtiqueta++;
                        }
                        if (pbCodBarras2.Visible)
                        {
                            QtdCodigoBarrasEtiqueta++;
                        }
                        if (pbCodBarras3.Visible)
                        {
                            QtdCodigoBarrasEtiqueta++;
                        }
                        if (pbCodBarras4.Visible)
                        {
                            QtdCodigoBarrasEtiqueta++;
                        }
                        if (pbCodBarras5.Visible)
                        {
                            QtdCodigoBarrasEtiqueta++;
                        }
                        if (pbCodBarras6.Visible)
                        {
                            QtdCodigoBarrasEtiqueta++;
                        }
                        if (ClienteLinha1 != null)
                        {
                            e.Graphics.DrawString(ClienteLinha1, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo);
                        }

                        if (ClienteLinha2 != null)
                        {
                            e.Graphics.DrawString(ClienteLinha2, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 12);
                        }

                        if (ClienteLinha3 != null)
                        {
                            e.Graphics.DrawString(ClienteLinha3, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 24);
                        }

                        if (ClienteLinha4 != null)
                        {
                            e.Graphics.DrawString(ClienteLinha4, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 36);
                        }

                        if (ClienteLinha5 != null)
                        {
                            e.Graphics.DrawString(ClienteLinha5, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 48);
                        }

                        if (ClienteLinha6 != null)
                        {
                            e.Graphics.DrawString(ClienteLinha6, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 60);
                        }

                        if (tbPedidoSM.Text != "")
                        {
                            e.Graphics.DrawString(tbPedidoSM.Text, fonteEmpresa, new SolidBrush(Color.Black), posicaoXPedidoSM, posicaoYPedidoSM + Intervalo);
                        }

                        if (tbCodigoSM.Text != "")
                        {
                            e.Graphics.DrawString(tbCodigoSM.Text, fonte, new SolidBrush(Color.Black), posicaoXCodigoSM, posicaoYCodigoSM + Intervalo);
                        }

                        if (tbCodMaterial.Text != "")
                        {
                            e.Graphics.DrawString(tbCodMaterial.Text, fonte, new SolidBrush(Color.Black), posicaoXCodigoMaterial, posicaoYCodigoMaterial + Intervalo);
                        }

                        if (tbPedido.Text != "")
                        {
                            e.Graphics.DrawString(tbPedido.Text, fonte, new SolidBrush(Color.Black), posicaoXPedido, posicaoYPedido + Intervalo);
                        }

                        if (DescricaoLinha1 != null)
                        {
                            e.Graphics.DrawString(DescricaoLinha1, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo);
                        }

                        if (DescricaoLinha2 != null)
                        {
                            e.Graphics.DrawString(DescricaoLinha2, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo + 12);
                        }

                        if (DescricaoLinha3 != null)
                        {
                            e.Graphics.DrawString(DescricaoLinha3, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo + 24);
                        }

                        if (DescricaoLinha4 != null)
                        {
                            e.Graphics.DrawString(DescricaoLinha4, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo + 36);
                        }

                        //if (tbDescricaoProduto.Text != "")
                        //{
                        //    e.Graphics.DrawString(tbDescricaoProduto.Text, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo);
                        //}

                        if (tbNotaFiscal.Text != "")
                        {
                            e.Graphics.DrawString(tbNotaFiscal.Text, fonte, new SolidBrush(Color.Black), posicaoXNotaFiscal, posicaoYNotaFiscal + Intervalo);
                        }

                        if (tbItem.Text != "")
                        {
                            e.Graphics.DrawString(tbItem.Text, fonte, new SolidBrush(Color.Black), posicaoXItem, posicaoYItem + Intervalo);
                        }

                        if (tbLote.Text != "")
                        {
                            e.Graphics.DrawString(tbLote.Text, fonte, new SolidBrush(Color.Black), posicaoXLote, posicaoYLote + Intervalo);
                        }

                        if (tbCaixa.Text != "")
                        {
                            e.Graphics.DrawString(tbCaixa.Text, fonte, new SolidBrush(Color.Black), posicaoXCaixa, posicaoYCaixa + Intervalo);
                        }

                        if (tbPesoLiquido.Text != "")
                        {
                            e.Graphics.DrawString((float.Parse(tbPesoLiquido.Text)*QtdCodigoBarrasEtiqueta).ToString("F" + 4), fonte, new SolidBrush(Color.Black), posicaoXPesoLiquido, posicaoYPesoLiquido + Intervalo);
                        }

                        if (tbPesoBruto.Text != "")
                        {
                            e.Graphics.DrawString((float.Parse(tbPesoBruto.Text)*QtdCodigoBarrasEtiqueta).ToString("F" + 4), fonte, new SolidBrush(Color.Black), posicaoXPesoBruto, posicaoYPesoBruto + Intervalo);
                        }

                        if (pbQRCode1.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode1.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 90, pbQRCode1.Image.Width, pbQRCode1.Image.Height);
                        }
                        if (pbCodBarras1.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras1.Image, (e.PageBounds.Right - 330 + pbQRCode1.Image.Width + 10), e.PageBounds.Top + Intervalo + 90, pbCodBarras1.Image.Width, pbCodBarras1.Image.Height);
                        }
                        udPages.DownButton();
                        Intervalo += 274;

                        if (s == QtdEtiquetas)
                        {
                            bTerminouImpressao = true;
                        }

                        if (((s % QtdEtiquetasPorPagina) == 0)&&(!bTerminouImpressao))
                        {
                            // Já imprimou 4 etiquetas
                            // Tem que trocar de página
                            PaginaAtualImpressao++;
                            e.HasMorePages = true;
                            break;
                        }
                    }
                }

                if (bTerminouImpressao)
                {
                    e.HasMorePages = false;
                }
            }
            else if(cbSelecao1Item.Text == "ETIQUETA ATUAL")
            {
                QtdCodigoBarrasEtiqueta = 0;
                if (pbCodBarras1.Visible)
                {
                    QtdCodigoBarrasEtiqueta++;
                }
                if (pbCodBarras2.Visible)
                {
                    QtdCodigoBarrasEtiqueta++;
                }
                if (pbCodBarras3.Visible)
                {
                    QtdCodigoBarrasEtiqueta++;
                }
                if (pbCodBarras4.Visible)
                {
                    QtdCodigoBarrasEtiqueta++;
                }
                if (pbCodBarras5.Visible)
                {
                    QtdCodigoBarrasEtiqueta++;
                }
                if (pbCodBarras6.Visible)
                {
                    QtdCodigoBarrasEtiqueta++;
                }
                if (ClienteLinha1 != null)
                {
                    e.Graphics.DrawString(ClienteLinha1, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo);
                }

                if (ClienteLinha2 != null)
                {
                    e.Graphics.DrawString(ClienteLinha2, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 12);
                }

                if (ClienteLinha3 != null)
                {
                    e.Graphics.DrawString(ClienteLinha3, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 24);
                }

                if (ClienteLinha4 != null)
                {
                    e.Graphics.DrawString(ClienteLinha4, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 36);
                }

                if (ClienteLinha5 != null)
                {
                    e.Graphics.DrawString(ClienteLinha5, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 48);
                }

                if (ClienteLinha6 != null)
                {
                    e.Graphics.DrawString(ClienteLinha6, fonteEmpresa, new SolidBrush(Color.Black), posicaoX, posicaoY + Intervalo + 60);
                }

                if (tbPedidoSM.Text != "")
                {
                    e.Graphics.DrawString(tbPedidoSM.Text, fonte, new SolidBrush(Color.Black), posicaoXPedidoSM, posicaoYPedidoSM + Intervalo);
                }

                if (tbCodigoSM.Text != "")
                {
                    e.Graphics.DrawString(tbCodigoSM.Text, fonte, new SolidBrush(Color.Black), posicaoXCodigoSM, posicaoYCodigoSM + Intervalo);
                }

                if (tbCodMaterial.Text != "")
                {
                    e.Graphics.DrawString(tbCodMaterial.Text, fonte, new SolidBrush(Color.Black), posicaoXCodigoMaterial, posicaoYCodigoMaterial + Intervalo);
                }

                if (tbPedido.Text != "")
                {
                    e.Graphics.DrawString(tbPedido.Text, fonte, new SolidBrush(Color.Black), posicaoXPedido, posicaoYPedido + Intervalo);
                }

                if (DescricaoLinha1 != null)
                {
                    e.Graphics.DrawString(DescricaoLinha1, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo);
                }

                if (DescricaoLinha2 != null)
                {
                    e.Graphics.DrawString(DescricaoLinha2, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo + 12);
                }

                if (DescricaoLinha3 != null)
                {
                    e.Graphics.DrawString(DescricaoLinha3, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo + 24);
                }

                if (DescricaoLinha4 != null)
                {
                    e.Graphics.DrawString(DescricaoLinha4, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo + 36);
                }

                //if (tbDescricaoProduto.Text != "")
                //{
                //    e.Graphics.DrawString(tbDescricaoProduto.Text, fonte, new SolidBrush(Color.Black), posicaoXDescricao, posicaoYDescricao + Intervalo);
                //}

                if (tbNotaFiscal.Text != "")
                {
                    e.Graphics.DrawString(tbNotaFiscal.Text, fonte, new SolidBrush(Color.Black), posicaoXNotaFiscal, posicaoYNotaFiscal + Intervalo);
                }

                if (tbItem.Text != "")
                {
                    e.Graphics.DrawString(tbItem.Text, fonte, new SolidBrush(Color.Black), posicaoXItem, posicaoYItem + Intervalo);
                }

                if (tbLote.Text != "")
                {
                    e.Graphics.DrawString(tbLote.Text, fonte, new SolidBrush(Color.Black), posicaoXLote, posicaoYLote + Intervalo);
                }

                if (tbCaixa.Text != "")
                {
                    e.Graphics.DrawString(tbCaixa.Text, fonte, new SolidBrush(Color.Black), posicaoXCaixa, posicaoYCaixa + Intervalo);
                }

                if (tbPesoLiquido.Text != "")
                {
                    e.Graphics.DrawString((float.Parse(tbPesoLiquido.Text)*QtdCodigoBarrasEtiqueta).ToString("F" + 4), fonte, new SolidBrush(Color.Black), posicaoXPesoLiquido, posicaoYPesoLiquido + Intervalo);
                }

                if (tbPesoBruto.Text != "")
                {
                    e.Graphics.DrawString((float.Parse(tbPesoBruto.Text)*QtdCodigoBarrasEtiqueta).ToString("F" + 4), fonte, new SolidBrush(Color.Black), posicaoXPesoBruto, posicaoYPesoBruto + Intervalo);
                }

                if (cbEtiquetaColetiva.Checked)
                {
                    if (QtdEtiquetasColetivas == 2)
                    {
                        if (pbQRCode1.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode1.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 90, pbQRCode1.Image.Width, pbQRCode1.Image.Height);
                        }
                        if (pbCodBarras1.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras1.Image, (e.PageBounds.Right - 330 + pbQRCode1.Image.Width + 10), e.PageBounds.Top + Intervalo + 90, pbCodBarras1.Image.Width, pbCodBarras1.Image.Height);
                        }
                        if (pbQRCode2.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode2.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 135, pbQRCode2.Image.Width, pbQRCode2.Image.Height);
                        }
                        if (pbCodBarras2.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras2.Image, (e.PageBounds.Right - 330 + pbQRCode2.Image.Width + 10), e.PageBounds.Top + Intervalo + 135, pbCodBarras2.Image.Width, pbCodBarras2.Image.Height);
                        }
                    }
                    else if (QtdEtiquetasColetivas == 3)
                    {
                        if (pbQRCode1.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode1.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 90, pbQRCode1.Image.Width, pbQRCode1.Image.Height);
                        }
                        if (pbCodBarras1.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras1.Image, (e.PageBounds.Right - 330 + pbQRCode1.Image.Width + 10), e.PageBounds.Top + Intervalo + 90, pbCodBarras1.Image.Width, pbCodBarras1.Image.Height);
                        }
                        if (pbQRCode2.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode2.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 135, pbQRCode2.Image.Width, pbQRCode2.Image.Height);
                        }
                        if (pbCodBarras2.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras2.Image, (e.PageBounds.Right - 330 + pbQRCode2.Image.Width + 10), e.PageBounds.Top + Intervalo + 135, pbCodBarras2.Image.Width, pbCodBarras2.Image.Height);
                        }
                        if (pbQRCode3.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode3.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 180, pbQRCode3.Image.Width, pbQRCode3.Image.Height);
                        }
                        if (pbCodBarras3.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras3.Image, (e.PageBounds.Right - 330 + pbQRCode3.Image.Width + 10), e.PageBounds.Top + Intervalo + 180, pbCodBarras3.Image.Width, pbCodBarras3.Image.Height);
                        }
                    }
                    else if (QtdEtiquetasColetivas == 4)
                    {
                        if (pbQRCode1.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode1.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 90, pbQRCode1.Image.Width, pbQRCode1.Image.Height);
                        }
                        if (pbCodBarras1.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras1.Image, (e.PageBounds.Right - 330 + pbQRCode1.Image.Width + 10), e.PageBounds.Top + Intervalo + 90, pbCodBarras1.Image.Width, pbCodBarras1.Image.Height);
                        }
                        if (pbQRCode2.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode2.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 135, pbQRCode2.Image.Width, pbQRCode2.Image.Height);
                        }
                        if (pbCodBarras2.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras2.Image, (e.PageBounds.Right - 330 + pbQRCode2.Image.Width + 10), e.PageBounds.Top + Intervalo + 135, pbCodBarras2.Image.Width, pbCodBarras2.Image.Height);
                        }
                        if (pbQRCode3.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode3.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 180, pbQRCode3.Image.Width, pbQRCode3.Image.Height);
                        }
                        if (pbCodBarras3.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras3.Image, (e.PageBounds.Right - 330 + pbQRCode3.Image.Width + 10), e.PageBounds.Top + Intervalo + 180, pbCodBarras3.Image.Width, pbCodBarras3.Image.Height);
                        }
                        if (pbQRCode4.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode4.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 225, pbQRCode4.Image.Width, pbQRCode4.Image.Height);
                        }
                        if (pbCodBarras4.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras4.Image, (e.PageBounds.Right - 330 + pbQRCode4.Image.Width + 10), e.PageBounds.Top + Intervalo + 225, pbCodBarras4.Image.Width, pbCodBarras4.Image.Height);
                        }
                    }
                    else if (QtdEtiquetasColetivas == 5)
                    {
                        if (pbQRCode1.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode1.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 90, pbQRCode1.Image.Width, pbQRCode1.Image.Height * 2 / 3);
                        }
                        if (pbCodBarras1.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras1.Image, (e.PageBounds.Right - 330 + pbQRCode1.Image.Width + 10), e.PageBounds.Top + Intervalo + 90, pbCodBarras1.Image.Width, pbCodBarras1.Image.Height * 2 / 3);
                        }
                        if (pbQRCode2.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode2.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 127, pbQRCode2.Image.Width, pbQRCode2.Image.Height * 2 / 3);
                        }
                        if (pbCodBarras2.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras2.Image, (e.PageBounds.Right - 330 + pbQRCode2.Image.Width + 10), e.PageBounds.Top + Intervalo + 127, pbCodBarras2.Image.Width, pbCodBarras2.Image.Height * 2 / 3);
                        }
                        if (pbQRCode3.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode3.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 144, pbQRCode3.Image.Width, pbQRCode3.Image.Height * 2 / 3);
                        }
                        if (pbCodBarras3.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras3.Image, (e.PageBounds.Right - 330 + pbQRCode3.Image.Width + 10), e.PageBounds.Top + Intervalo + 144, pbCodBarras3.Image.Width, pbCodBarras3.Image.Height * 2 / 3);
                        }
                        if (pbQRCode4.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode4.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 171, pbQRCode4.Image.Width, pbQRCode4.Image.Height * 2 / 3);
                        }
                        if (pbCodBarras4.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras4.Image, (e.PageBounds.Right - 330 + pbQRCode4.Image.Width + 10), e.PageBounds.Top + Intervalo + 171, pbCodBarras4.Image.Width, pbCodBarras4.Image.Height * 2 / 3);
                        }
                        if (pbQRCode5.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode5.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 198, pbQRCode5.Image.Width, pbQRCode5.Image.Height * 2 / 3);
                        }
                        if (pbCodBarras5.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras5.Image, (e.PageBounds.Right - 330 + pbQRCode5.Image.Width + 10), e.PageBounds.Top + Intervalo + 198, pbCodBarras5.Image.Width, pbCodBarras5.Image.Height * 2 / 3);
                        }
                    }
                    else if(QtdEtiquetasColetivas == 6)
                    {
                        if (pbQRCode1.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode1.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 90, pbQRCode1.Image.Width, pbQRCode1.Image.Height*2/3);
                        }
                        if (pbCodBarras1.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras1.Image, (e.PageBounds.Right - 330 + pbQRCode1.Image.Width + 10), e.PageBounds.Top + Intervalo + 90, pbCodBarras1.Image.Width, pbCodBarras1.Image.Height*2/3);
                        }
                        if (pbQRCode2.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode2.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 127, pbQRCode2.Image.Width, pbQRCode2.Image.Height*2/3);
                        }
                        if (pbCodBarras2.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras2.Image, (e.PageBounds.Right - 330 + pbQRCode2.Image.Width + 10), e.PageBounds.Top + Intervalo + 127, pbCodBarras2.Image.Width, pbCodBarras2.Image.Height*2/3);
                        }
                        if (pbQRCode3.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode3.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 144, pbQRCode3.Image.Width, pbQRCode3.Image.Height*2/3);
                        }
                        if (pbCodBarras3.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras3.Image, (e.PageBounds.Right - 330 + pbQRCode3.Image.Width + 10), e.PageBounds.Top + Intervalo + 144, pbCodBarras3.Image.Width, pbCodBarras3.Image.Height*2/3);
                        }
                        if (pbQRCode4.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode4.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 171, pbQRCode4.Image.Width, pbQRCode4.Image.Height*2/3);
                        }
                        if (pbCodBarras4.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras4.Image, (e.PageBounds.Right - 330 + pbQRCode4.Image.Width + 10), e.PageBounds.Top + Intervalo + 171, pbCodBarras4.Image.Width, pbCodBarras4.Image.Height*2/3);
                        }
                        if (pbQRCode5.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode5.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 198, pbQRCode5.Image.Width, pbQRCode5.Image.Height*2/3);
                        }
                        if (pbCodBarras5.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras5.Image, (e.PageBounds.Right - 330 + pbQRCode5.Image.Width + 10), e.PageBounds.Top + Intervalo + 198, pbCodBarras5.Image.Width, pbCodBarras5.Image.Height*2/3);
                        }
                        if (pbQRCode6.Visible)
                        {
                            e.Graphics.DrawImage(pbQRCode6.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 225, pbQRCode6.Image.Width, pbQRCode6.Image.Height*2/3);
                        }
                        if (pbCodBarras6.Visible)
                        {
                            e.Graphics.DrawImage(pbCodBarras6.Image, (e.PageBounds.Right - 330 + pbQRCode6.Image.Width + 10), e.PageBounds.Top + Intervalo + 225, pbCodBarras6.Image.Width, pbCodBarras6.Image.Height*2/3);
                        }
                    }
                }
                else
                {
                    // Etiqueta Individual
                    if (pbQRCode1.Visible)
                    {
                        e.Graphics.DrawImage(pbQRCode1.Image, e.PageBounds.Right - 330, e.PageBounds.Top + Intervalo + 90, pbQRCode1.Image.Width, pbQRCode1.Image.Height);
                    }
                    if (pbCodBarras1.Visible)
                    {
                        e.Graphics.DrawImage(pbCodBarras1.Image, (e.PageBounds.Right - 330 + pbQRCode1.Image.Width + 10), e.PageBounds.Top + Intervalo + 90, pbCodBarras1.Image.Width, pbCodBarras1.Image.Height);
                    }
                }
                udPages.DownButton();
                Intervalo += 274;
            }
        }

        private void pbCodBarras1_Click(object sender, EventArgs e)
        {

        }

        private void pbCodBarras2_Click(object sender, EventArgs e)
        {

        }

        private void panelNroSeries_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pbQRCode1_Click(object sender, EventArgs e)
        {

        }

        private void cbSelecao1Item_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void cbSelecao1Item_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void cbTipoCodBarras_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void udPages_SelectedItemChanged(object sender, EventArgs e)
        {
            String[] BarCodeSplit;
            String BarCodeStr;
            int IndiceAux;

            if ((PaginaSelecionada.ToString() == udPages.Text)||(udPages.Text == ""))
            {
                return;
            }

            PaginaSelecionada = int.Parse(udPages.Text);

            if (cbEtiquetaColetiva.Checked)
            {
                if (cbPedidoParcial.Checked)
                {
                    if (!cbEtiquetaColetiva.Checked)
                    {
                        tbCaixa.Text = ((((int.Parse(tbNroInicial.Text) - 1) * int.Parse(tbQtdColetiva.Text)) + 1) + PaginaSelecionada - 1).ToString() + "/" + clGlobal.QtdEtiquetasIndividuais.ToString();
                    }
                    else
                    {
                        tbCaixa.Text = (PaginaSelecionada + int.Parse(tbNroInicial.Text) - 1).ToString() + "/" + clGlobal.QtdEtiquetasColetivas;
                    }
                }
                else
                {
                    tbCaixa.Text = PaginaSelecionada.ToString() + "/" + QtdEtiquetas;
                }
                //tbCaixa.Text = PaginaSelecionada.ToString() + "/" + QtdEtiquetas;
            }
            else if(cbEtiquetaIndividual.Checked)
            {
                if (cbPedidoParcial.Checked)
                {
                    if (!cbEtiquetaColetiva.Checked)
                    {
                        tbCaixa.Text = (int.Parse(tbNroInicial.Text) + PaginaSelecionada - 1).ToString() + "/" + clGlobal.QtdEtiquetasIndividuais.ToString();
                    }
                    else
                    {
                        tbCaixa.Text = (PaginaSelecionada + (int.Parse(tbNroInicial.Text) / QtdEtiquetasColetivas)).ToString() + "/" + clGlobal.QtdEtiquetasColetivas;
                    }
                }
                else
                {
                    tbCaixa.Text = PaginaSelecionada.ToString() + "/" + QtdEtiquetas;
                }
                //tbCaixa.Text = PaginaSelecionada.ToString() + "/" + QtdEtiquetas;
            }

            //if (cbEtiquetaIndividual.Checked)
            //{
            //    if (PaginaSelecionada > 60)
            //    {
            //        //tbCaixa.Text = PaginaSelecionada.ToString() + "/" + QtdEtiquetas;
            //        tbCaixa.Text = (PaginaSelecionada - 60).ToString() + "/40";
            //    }
            //    else
            //    {
            //        tbCaixa.Text = "";
            //    }
            //}
            //else if (cbEtiquetaColetiva.Checked)
            //{
            //    if (PaginaSelecionada > 15)
            //    {
            //        //tbCaixa.Text = PaginaSelecionada.ToString() + "/" + QtdEtiquetas;
            //        tbCaixa.Text = (PaginaSelecionada - 15).ToString() + "/10";
            //    }
            //    else
            //    {
            //        tbCaixa.Text = "";
            //    }
            //}

            // Colocar na tela os numeros de série
            if (ListaNroSerie != null)
            {
                if (cbEtiquetaColetiva.Checked)
                {
                    if (QtdEtiquetasColetivas == 2)
                    {
                        IndiceAux = 0 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode1.Image = GerarQRCode(pbQRCode1.Width, pbQRCode1.Height, BarCodeStr);
                            pbCodBarras1.Image = GerarBarCode(pbCodBarras1.Width, pbCodBarras1.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode1.Visible = true;
                            pbCodBarras1.Visible = true;
                        }
                        else
                        {
                            pbQRCode1.Visible = false;
                            pbCodBarras1.Visible = false;
                        }

                        IndiceAux = 1 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode2.Image = GerarQRCode(pbQRCode2.Width, pbQRCode2.Height, BarCodeStr);
                            pbCodBarras2.Image = GerarBarCode(pbCodBarras2.Width, pbCodBarras2.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode2.Visible = true;
                            pbCodBarras2.Visible = true;
                        }
                        else
                        {
                            pbQRCode2.Visible = false;
                            pbCodBarras2.Visible = false;
                        }
                    }
                    else if (QtdEtiquetasColetivas == 3)
                    {
                        IndiceAux = 0 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode1.Image = GerarQRCode(pbQRCode1.Width, pbQRCode1.Height, BarCodeStr);
                            pbCodBarras1.Image = GerarBarCode(pbCodBarras1.Width, pbCodBarras1.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode1.Visible = true;
                            pbCodBarras1.Visible = true;
                        }
                        else
                        {
                            pbQRCode1.Visible = false;
                            pbCodBarras1.Visible = false;
                        }

                        IndiceAux = 1 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode2.Image = GerarQRCode(pbQRCode2.Width, pbQRCode2.Height, BarCodeStr);
                            pbCodBarras2.Image = GerarBarCode(pbCodBarras2.Width, pbCodBarras2.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode2.Visible = true;
                            pbCodBarras2.Visible = true;
                        }
                        else
                        {
                            pbQRCode2.Visible = false;
                            pbCodBarras2.Visible = false;
                        }

                        IndiceAux = 2 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode3.Image = GerarQRCode(pbQRCode3.Width, pbQRCode3.Height, BarCodeStr);
                            pbCodBarras3.Image = GerarBarCode(pbCodBarras3.Width, pbCodBarras3.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode3.Visible = true;
                            pbCodBarras3.Visible = true;
                        }
                        else
                        {
                            pbQRCode3.Visible = false;
                            pbCodBarras3.Visible = false;
                        }
                    }
                    else if (QtdEtiquetasColetivas == 4)
                    {
                        IndiceAux = 0 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode1.Image = GerarQRCode(pbQRCode1.Width, pbQRCode1.Height, BarCodeStr);
                            pbCodBarras1.Image = GerarBarCode(pbCodBarras1.Width, pbCodBarras1.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode1.Visible = true;
                            pbCodBarras1.Visible = true;
                        }
                        else
                        {
                            pbQRCode1.Visible = false;
                            pbCodBarras1.Visible = false;
                        }

                        IndiceAux = 1 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode2.Image = GerarQRCode(pbQRCode2.Width, pbQRCode2.Height, BarCodeStr);
                            pbCodBarras2.Image = GerarBarCode(pbCodBarras2.Width, pbCodBarras2.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode2.Visible = true;
                            pbCodBarras2.Visible = true;
                        }
                        else
                        {
                            pbQRCode2.Visible = false;
                            pbCodBarras2.Visible = false;
                        }

                        IndiceAux = 2 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode3.Image = GerarQRCode(pbQRCode3.Width, pbQRCode3.Height, BarCodeStr);
                            pbCodBarras3.Image = GerarBarCode(pbCodBarras3.Width, pbCodBarras3.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode3.Visible = true;
                            pbCodBarras3.Visible = true;
                        }
                        else
                        {
                            pbQRCode3.Visible = false;
                            pbCodBarras3.Visible = false;
                        }

                        IndiceAux = 3 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode4.Image = GerarQRCode(pbQRCode4.Width, pbQRCode4.Height, BarCodeStr);
                            pbCodBarras4.Image = GerarBarCode(pbCodBarras4.Width, pbCodBarras4.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode4.Visible = true;
                            pbCodBarras4.Visible = true;
                        }
                        else
                        {
                            pbQRCode4.Visible = false;
                            pbCodBarras4.Visible = false;
                        }
                    }
                    else if (QtdEtiquetasColetivas == 5)
                    {
                        IndiceAux = 0 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode1.Image = GerarQRCode(pbQRCode1.Width, pbQRCode1.Height, BarCodeStr);
                            pbCodBarras1.Image = GerarBarCode(pbCodBarras1.Width, pbCodBarras1.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode1.Visible = true;
                            pbCodBarras1.Visible = true;
                        }
                        else
                        {
                            pbQRCode1.Visible = false;
                            pbCodBarras1.Visible = false;
                        }

                        IndiceAux = 1 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode2.Image = GerarQRCode(pbQRCode2.Width, pbQRCode2.Height, BarCodeStr);
                            pbCodBarras2.Image = GerarBarCode(pbCodBarras2.Width, pbCodBarras2.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode2.Visible = true;
                            pbCodBarras2.Visible = true;
                        }
                        else
                        {
                            pbQRCode2.Visible = false;
                            pbCodBarras2.Visible = false;
                        }

                        IndiceAux = 2 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode3.Image = GerarQRCode(pbQRCode3.Width, pbQRCode3.Height, BarCodeStr);
                            pbCodBarras3.Image = GerarBarCode(pbCodBarras3.Width, pbCodBarras3.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode3.Visible = true;
                            pbCodBarras3.Visible = true;
                        }
                        else
                        {
                            pbQRCode3.Visible = false;
                            pbCodBarras3.Visible = false;
                        }

                        IndiceAux = 3 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode4.Image = GerarQRCode(pbQRCode4.Width, pbQRCode4.Height, BarCodeStr);
                            pbCodBarras4.Image = GerarBarCode(pbCodBarras4.Width, pbCodBarras4.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode4.Visible = true;
                            pbCodBarras4.Visible = true;
                        }
                        else
                        {
                            pbQRCode4.Visible = false;
                            pbCodBarras4.Visible = false;
                        }

                        IndiceAux = 4 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode5.Image = GerarQRCode(pbQRCode5.Width, pbQRCode5.Height, BarCodeStr);
                            pbCodBarras5.Image = GerarBarCode(pbCodBarras5.Width, pbCodBarras5.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode5.Visible = true;
                            pbCodBarras5.Visible = true;
                        }
                        else
                        {
                            pbQRCode5.Visible = false;
                            pbCodBarras5.Visible = false;
                        }
                    }
                    else if(QtdEtiquetasColetivas == 6)
                    {
                        IndiceAux = 0 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode1.Image = GerarQRCode(pbQRCode1.Width, pbQRCode1.Height, BarCodeStr);
                            pbCodBarras1.Image = GerarBarCode(pbCodBarras1.Width, pbCodBarras1.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode1.Visible = true;
                            pbCodBarras1.Visible = true;
                        }
                        else
                        {
                            pbQRCode1.Visible = false;
                            pbCodBarras1.Visible = false;
                        }

                        IndiceAux = 1 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode2.Image = GerarQRCode(pbQRCode2.Width, pbQRCode2.Height, BarCodeStr);
                            pbCodBarras2.Image = GerarBarCode(pbCodBarras2.Width, pbCodBarras2.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode2.Visible = true;
                            pbCodBarras2.Visible = true;
                        }
                        else
                        {
                            pbQRCode2.Visible = false;
                            pbCodBarras2.Visible = false;
                        }

                        IndiceAux = 2 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode3.Image = GerarQRCode(pbQRCode3.Width, pbQRCode3.Height, BarCodeStr);
                            pbCodBarras3.Image = GerarBarCode(pbCodBarras3.Width, pbCodBarras3.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode3.Visible = true;
                            pbCodBarras3.Visible = true;
                        }
                        else
                        {
                            pbQRCode3.Visible = false;
                            pbCodBarras3.Visible = false;
                        }

                        IndiceAux = 3 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode4.Image = GerarQRCode(pbQRCode4.Width, pbQRCode4.Height, BarCodeStr);
                            pbCodBarras4.Image = GerarBarCode(pbCodBarras4.Width, pbCodBarras4.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode4.Visible = true;
                            pbCodBarras4.Visible = true;
                        }
                        else
                        {
                            pbQRCode4.Visible = false;
                            pbCodBarras4.Visible = false;
                        }

                        IndiceAux = 4 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode5.Image = GerarQRCode(pbQRCode5.Width, pbQRCode5.Height, BarCodeStr);
                            pbCodBarras5.Image = GerarBarCode(pbCodBarras5.Width, pbCodBarras5.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode5.Visible = true;
                            pbCodBarras5.Visible = true;
                        }
                        else
                        {
                            pbQRCode5.Visible = false;
                            pbCodBarras5.Visible = false;
                        }

                        IndiceAux = 5 + ((PaginaSelecionada - 1) * QtdEtiquetasColetivas);
                        if (IndiceAux < QtdItens)
                        {
                            BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                            if (BarCodeSplit.Length > 1)
                            {
                                BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                            }
                            else
                            {
                                BarCodeStr = ListaNroSerie[IndiceAux];
                            }

                            pbQRCode6.Image = GerarQRCode(pbQRCode6.Width, pbQRCode6.Height, BarCodeStr);
                            pbCodBarras6.Image = GerarBarCode(pbCodBarras6.Width, pbCodBarras6.Height, BarCodeStr, cbTipoCodBarras.Text);
                            //pbQRCode6.Visible = true;
                            pbCodBarras6.Visible = true;
                        }
                        else
                        {
                            pbQRCode6.Visible = false;
                            pbCodBarras6.Visible = false;
                        }
                    }
                }
                else
                {
                    // Etiqueta individual
                    IndiceAux = PaginaSelecionada - 1;
                    if (IndiceAux < QtdItens)
                    {
                        BarCodeSplit = ListaNroSerie[IndiceAux].Split('-');
                        if (BarCodeSplit.Length > 1)
                        {
                            BarCodeStr = BarCodeSplit[0] + BarCodeSplit[1];
                        }
                        else
                        {
                            BarCodeStr = ListaNroSerie[IndiceAux];
                        }

                        pbQRCode1.Image = GerarQRCode(pbQRCode1.Width, pbQRCode1.Height, BarCodeStr);
                        pbCodBarras1.Image = GerarBarCode(pbCodBarras1.Width, pbCodBarras1.Height, BarCodeStr, cbTipoCodBarras.Text);
                        //pbQRCode1.Visible = true;
                        pbCodBarras1.Visible = true;
                    }
                    else
                    {
                        pbQRCode1.Visible = false;
                        pbCodBarras1.Visible = false;
                    }
                }
            }
        }

        private void visualizarImpressãoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printDocument1.DocumentName = "Etiqueta MGE";
            Font f = new Font("Courier New", 10);
            ////  Variável para armazenamento de posicao vertical.
            int posY = printDocument1.DefaultPageSettings.Margins.Top;
            printDocument1.DefaultPageSettings.Landscape = false;

            printPreviewDialog1.Document = printDocument1;
            printPreviewDialog1.ShowDialog();
        }

        private void cbEtiquetaColetiva_CheckedChanged(object sender, EventArgs e)
        {
            if (cbEtiquetaColetiva.Checked)
            {
                cbEtiquetaIndividual.Checked = false;
            }
            else
            {
                cbEtiquetaIndividual.Checked = true;
            }

            if ((tbPedidoSM.Text != "")&&(tbPedido.Text != ""))
            {
                button1_Click(sender, e);
            }
        }

        private void cbEtiquetaIndividual_CheckedChanged(object sender, EventArgs e)
        {
            if (cbEtiquetaIndividual.Checked)
            {
                cbEtiquetaColetiva.Checked = false;
            }
            else
            {
                cbEtiquetaColetiva.Checked = true;
            }

            if ((tbPedidoSM.Text != "") && (tbPedido.Text != ""))
            {
                button1_Click(sender, e);
            }
        }

        private void cbPedidoParcial_CheckedChanged(object sender, EventArgs e)
        {
            if (cbPedidoParcial.Checked)
            {
                lblNroInicial.Visible = true;
                tbNroInicial.Visible = true;
                tbNroInicial.Enabled = true;
                lblQtdItensImprimir.Visible = true;
                tbQtdItensImprimir.Visible = true;
                tbQtdItensImprimir.Enabled = true;
            }
            else
            {
                lblNroInicial.Visible = false;
                tbNroInicial.Visible = false;
                tbNroInicial.Enabled = false;
                lblQtdItensImprimir.Visible = false;
                tbQtdItensImprimir.Visible = false;
                tbQtdItensImprimir.Enabled = false;
            }
        }

        private void tbNroInicial_TextChanged(object sender, EventArgs e)
        {

        }

        private void tbQtdColetiva_TextChanged(object sender, EventArgs e)
        {
            QtdEtiquetasColetivas = int.Parse(tbQtdColetiva.Text);
        }

        public static async Task TestaPing(string url)
        {
            IPHostEntry Host;
            try
            {
                Ping Pinger = new Ping();
                PingReply resposta = await Pinger.SendPingAsync(url);
                Host = Dns.GetHostByAddress(IPAddress.Parse(url));
                clGlobal.HostName = Host.HostName.ToString();

                ExibeRespostaPing(resposta);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }            
        }

        private static void ExibeRespostaPing(PingReply resposta)
        {
            if(resposta.Status == IPStatus.Success)
            {
                clGlobal.bRespostaPing = true;
            }
            else
            {
                clGlobal.bRespostaPing = false;
            }

            clGlobal.bTerminouPing = true;
        }

        private void tbOrdemProducao_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                string connString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" + clGlobal.EnderecoServidor + ")(PORT=1521))(CONNECT_DATA = (SERVER = DEDICATED)(SERVICE_NAME = XE)))";
                //string connString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=mgers.dyndns.org)(PORT=1521))(CONNECT_DATA = (SERVER = DEDICATED)(SERVICE_NAME = XE)))";
                connString = connString + ";User Id=ODIN_CONSULTA;Password=ODIN123@;";
                OracleConnection conn = new OracleConnection(connString);
                conn.Open();
                if (conn.State != ConnectionState.Open)
                {
                    MessageBox.Show("Não foi possível a conexão com o banco ODIN", "Erro");
                    return;
                }

                OracleCommand commandOrdemProducao = conn.CreateCommand();
                string sqlOrdemProducao = "SELECT * FROM ODIN_MGE.ordem_producao where nro_ordem_producao = " + tbOrdemProducao.Text;
                commandOrdemProducao.CommandText = sqlOrdemProducao;

                OracleDataReader readerOrdemProducao = commandOrdemProducao.ExecuteReader();

                if (readerOrdemProducao.HasRows)
                {
                    while (readerOrdemProducao.Read())
                    {
                        if ((readerOrdemProducao["NRO_PEDIDO"] != null))
                        {
                            tbPedidoOdin.Text = readerOrdemProducao["NRO_PEDIDO"].ToString();
                            if(tbPedidoOdin.Text != "")
                            {
                                button1_Click(sender, e);
                            }
                            else
                            {
                                tbPedidoOdin.Focus();
                            }
                        }
                        else
                        {
                            tbPedidoOdin.Focus();
                        }
                    }
                }
                else
                {
                    tbPedidoOdin.Focus();
                }

                conn.Close();
            }
        }

        private void tbPedidoOdin_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                button1_Click(sender, e);
            }
        }
    }
}
