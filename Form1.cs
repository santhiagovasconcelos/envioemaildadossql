using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using ND_Wartsila.Classes;

using System.Data.OleDb;

namespace ND_Wartsila
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            /*
            string ano = Convert.ToString(DateTime.Now.Year);
            string mes = Convert.ToString(DateTime.Now.Month);

            comboBoxAno.Text = ano;
            comboBoxMes.Text = mes;*/

        }

     


        public void CarregarBancoTotvs(string mes, string ano)
        {
            //set the connection string
            string connString = @"Data Source=192.168.0.2;Initial Catalog=TOTVS;User ID=sa;Password=SenhaAqui.123";

            string dataCompleta = ano + mes;

            //checando se os campos do intervalo da nota estão vazios para prosseguir
            if (comboBoxMes.Text == "" | comboBoxAno.Text == "")
            {
                MessageBox.Show("Favor inserir data para prosseguir");
                return;
            }
            else
            {
                try
                {
                    //sql connection object
                    using (SqlConnection conn = new SqlConnection(connString))
                    {

                        //retrieve the SQL Server instance version
                        string query = @"SELECT C7_FORNECE, C7_OBS, C7_NUM, C7_TOTAL FROM SC7010 WHERE C7_FORNECE = '000131' AND C7_EMISSAO LIKE '" + dataCompleta + "%' OR C7_FORNECE = '001562' AND C7_EMISSAO LIKE '" + dataCompleta + "%' ORDER BY C7_OBS;";

                        //define the SqlCommand object
                        SqlCommand cmd = new SqlCommand(query, conn);


                        //Set the SqlDataAdapter object
                        SqlDataAdapter dAdapter = new SqlDataAdapter(cmd);

                        //define dataset
                        DataSet ds = new DataSet();

                        //fill dataset with query results
                        dAdapter.Fill(ds);

                        //set the DataGridView control's data source/data table
                        dataGridView1.DataSource = ds.Tables[0];

                        //close connection
                        conn.Close();
                    }

                    //renomeando as colunas e colocando autosize
                    dataGridView1.Columns[0].HeaderText = "FORNECEDOR";
                    dataGridView1.Columns[1].HeaderText = "DESCRIÇÃO";
                    dataGridView1.Columns[2].HeaderText = "NUM PEDIDO";
                    dataGridView1.Columns[3].HeaderText = "VALOR";

                    //corrigindo os dados dentro da coluna
                    dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

                    //adicionando coluna para receber o caminho
                    dataGridView1.Columns.Add("caminho", "CAMINHO DO ARQUIVO");
                    dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

                    //SALVANDO EXCEL e populando grid com o caminho
                    try
                    {
                        string dataNoExcel = comboBoxMes.Text + "/" + comboBoxAno.Text;
                        string fixoClaro = Convert.ToString(dataGridView1.Rows[0].Cells[3].Value);
                        string dadosClaro = Convert.ToString(dataGridView1.Rows[1].Cells[3].Value);
                        string dadosWorldnet = Convert.ToString(dataGridView1.Rows[2].Cells[3].Value);

                        SalvarExcel(dataNoExcel, fixoClaro, dadosClaro, dadosWorldnet);

                        int qtdLinha = dataGridView1.Rows.Count - 1;
                        int cont = 0;

                        while (cont != qtdLinha)
                        {
                            string fornecedor = Convert.ToString(dataGridView1.Rows[cont].Cells[0].Value);
                            string pedido = Convert.ToString(dataGridView1.Rows[cont].Cells[2].Value);

                            AdicionarCaminhoGrid(fornecedor, pedido, cont);

                            cont = cont + 1;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Erro ao salvar excel ou popular caminho dos arquivos");
                        return;
                    }

                }
                catch (Exception ex)
                {
                    //display error message
                    MessageBox.Show("Favor verificar o intervalo da numeração das notas com os ZEROS na frente do número!");
                    MessageBox.Show("Exception: " + ex.Message);
                }

                

            }
        }

        //ADICIONANDO A COLUNA DO CAMINHO DO ARQUIVO
        public void AdicionarCaminhoGrid(string fornecedor, string pedido, int linha)
        {
            string connString = @"Data Source=192.168.0.2;Initial Catalog=TOTVS;User ID=sa;Password=SenhaAqui.123";

            try
            {
                //sql connection object
                using (SqlConnection conn = new SqlConnection(connString))
                {

                    //retrieve the SQL Server instance version
                    string query = @"SELECT ZS0_ARQUIV FROM ZS0010 where ZS0_FORNEC = '" + fornecedor + "' and ZS0_PEDIDO = '" + pedido + "';";

                    //define the SqlCommand object
                    SqlCommand cmd = new SqlCommand(query, conn);

                    conn.Open();

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            //string caminho = "\\\\192.168.0.3\\entrada\\" + String.Format("{0}", reader[0]);
                            string caminho = String.Format("{0}", reader[0]);

                            dataGridView1.Rows[linha].Cells[4].Value = caminho;
                            
                        }
                    }

                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                //display error message
                MessageBox.Show("Favor verificar o intervalo da numeração das notas com os ZEROS na frente do número!");
                MessageBox.Show("Exception: " + ex.Message);
            }
        }


        //CARREGAR DADOS
        private void button1_Click(object sender, EventArgs e)
        {
            CarregarBancoTotvs(comboBoxMes.Text, comboBoxAno.Text);
            comboBoxAno.Enabled = false;
            comboBoxMes.Enabled = false;
            button2.Enabled = true;
            button1.Enabled = false;
            button3.Enabled = true;
        }

        //criar excel
        private void button3_Click(object sender, EventArgs e)
        {
            string dataEnvio = comboBoxMes.Text + "-" + comboBoxAno.Text;

            ND_Wartsila.Classes.Email enviar = new ND_Wartsila.Classes.Email();

            enviar.EnviarND(textPara.Text, textLogin.Text, textSenha.Text, Convert.ToString(dataGridView1.Rows[0].Cells[4].Value), Convert.ToString(dataGridView1.Rows[1].Cells[4].Value), Convert.ToString(dataGridView1.Rows[2].Cells[4].Value), dataEnvio, Convert.ToString(dataGridView1.Rows[0].Cells[3].Value), Convert.ToString(dataGridView1.Rows[1].Cells[3].Value), Convert.ToString(dataGridView1.Rows[2].Cells[3].Value), Convert.ToString(dataGridView1.Rows[0].Cells[1].Value), Convert.ToString(dataGridView1.Rows[1].Cells[1].Value), Convert.ToString(dataGridView1.Rows[2].Cells[1].Value));
            //string emailPara, string emailLogin, string emailSenha, string caminhoClaroFixo, string caminhoClaroDados, string caminhoWorldnetDados, string data, string valorClaroFixo, string valorClaroDados, string valorWoldnetDados, string descricaoClaroFixo, string descricaoClaroDados, string descricaoWorldnetDados)
        }

        public void SalvarExcel(string data, string fixoClaro, string dadosClaro, string dadosWorldnet)
        {
            try
            {
                using var wbook = new XLWorkbook("C:\\ND_WARTSILA\\Detalhamento.xlsx");

                var ws = wbook.Worksheet(1);

                ws.Cell(5, 4).Value = data;
                ws.Cell(8, 4).Value = fixoClaro;
                ws.Cell(11, 4).Value = dadosClaro;
                ws.Cell(14, 4).Value = dadosWorldnet;

                wbook.Save();
            }
            catch
            {
                MessageBox.Show("Erro ao salvar excel. Favor verificar se o mesmo já não está aberto!");
            }

            //https://zetcode.com/csharp/excel/
        }
		/*
        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //int posicao = dataGridView1.SelectedRows.

            string caminho = Convert.ToString(dataGridView1["CAMINHO", e.RowIndex].Value);

            MessageBox.Show(caminho);
			
            System.Diagnostics.Process.Start(@"\\192.168.0.3\entrada\0000000056160950-E-161023.xlsx");
        }*/

        private void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            button1.Enabled = true;
            button3.Enabled = false;

            comboBoxAno.Text = "";
            comboBoxMes.Text = "";
            comboBoxAno.Enabled = true;
            comboBoxMes.Enabled = true;

        }
    }
}
