using System.Collections;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Xml.Linq;
using System;
using ClosedXML.Excel;
using System.Linq;
using ClosedXML.Report.Utils;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Bibliography;
using System.Threading;
using System.Threading.Tasks;

namespace friendly_bpa2excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            label1.Text = "BPA";
            textBox1.Location = new Point(68, 34);
            textBox1.Size = new Size(264, 23);
            label1.Visible = true;
            textBox1.Visible = true;
            button3.Visible = true;
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!File.Exists("bpai.txt") || !File.Exists("bpac.txt"))
            {
                MessageBox.Show("Algum dos arquivos não foi encontrado", "Falha", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                label1.Visible = true;
                textBox1.Visible = true;
                button3.Visible = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "D: \\Users\\rarodrigues\\Desktop\\codes\\c#\\Teste\\bin\\Debug\\net6.0-windows\\";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "D: \\Users\\rarodrigues\\Desktop\\codes\\c#\\Teste\\bin\\Debug\\net6.0-windows\\";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialog1.FileName;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "D: \\Users\\rarodrigues\\Desktop\\codes\\c#\\Teste\\bin\\Debug\\net6.0-windows\\";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = openFileDialog1.FileName;
            }
        }

        private async Task proc(string file_out, string file_c, string file_i, ProgressBar pb)
        {
            new BPABuilder(file_out, file_c, file_i, pb);
            MessageBox.Show("Arquivo salvo em " + file_out, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }

        public void pos_proc()
        {
            Console.Write("pos_proc started...");
            progressBar1.Visible = false;
        }

        private async void middle(string file_out, string file_c, string file_i, ProgressBar pb)
        {

            await proc(file_out, file_c, file_i, pb);
            textBox1.Invoke((Action)delegate () { pos_proc(); });
        }

        private void convert()
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
            saveFileDialog1.RestoreDirectory = true;
            progressBar1.Visible = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Thread t1 = new Thread(new ThreadStart((Action)delegate ()
                {
                    middle(saveFileDialog1.FileName, textBox2.Text, textBox3.Text, progressBar1);

                }));

                t1.Start();

            }

        }

        private void button4_Click(object sender, EventArgs e)
        {

            if (checkBox1.Checked)
            {
                progressBar1.Visible = true;
                BPA bpa = new BPA(textBox1.Text);
                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    bpa.dropBPA2Excel(folderBrowserDialog.SelectedPath + "\\bpac.txt", folderBrowserDialog.SelectedPath + "\\bpai.txt", progressBar1);
                    MessageBox.Show("Arquivos salvos em " + folderBrowserDialog.SelectedPath, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                progressBar1.Visible = false;
            }
            else if (checkBox2.Checked)
            {
                convert();
            }

        }

        private void textBox1_textChange(object sender, EventArgs e)
        {
            if (textBox1.Text.Length == 0 || textBox1.Text.LastIndexOf(".") == -1 || textBox1.Text.LastIndexOf(".") < textBox1.Text.LastIndexOf("\\") || !File.Exists(textBox1.Text))
            {
                button4.Visible = false;
            }
            else
            {
                button4.Visible = true;
            }
        }

        private void textBox2_textChange(object sender, EventArgs e)
        {
            if (textBox2.Text.Length == 0 || textBox2.Text.LastIndexOf(".") == -1 || textBox2.Text.LastIndexOf(".") < textBox2.Text.LastIndexOf("\\") || !File.Exists(textBox2.Text) || textBox3.Text.Length == 0 || textBox3.Text.LastIndexOf(".") == -1 || textBox3.Text.LastIndexOf(".") < textBox3.Text.LastIndexOf("\\") || !File.Exists(textBox3.Text))
            {
                button4.Visible = false;
            }
            else
            {
                button4.Visible = true;
            }
        }

        private void textBox3_textChange(object sender, EventArgs e)
        {
            if (textBox2.Text.Length == 0 || textBox2.Text.LastIndexOf(".") == -1 || textBox2.Text.LastIndexOf(".") < textBox2.Text.LastIndexOf("\\") || !File.Exists(textBox2.Text) || textBox3.Text.Length == 0 || textBox3.Text.LastIndexOf(".") == -1 || textBox3.Text.LastIndexOf(".") < textBox3.Text.LastIndexOf("\\") || !File.Exists(textBox3.Text))
            {
                button4.Visible = false;
            }
            else
            {
                button4.Visible = true;
            }
        }

        

        private void Form1_Load(object sender, EventArgs e)
        {
            //Text = GetNextCellName("ZZ");
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                if (checkBox2.Checked)
                {
                    checkBox2.Checked = false;
                }

                label1.Text = "BPA";
                label1.Visible = true;
                textBox1.Visible = true;
                button3.Visible = true;

                if (textBox1.Text.Length == 0 || textBox1.Text.LastIndexOf(".") == -1 || textBox1.Text.LastIndexOf(".") < textBox1.Text.LastIndexOf("\\") || !File.Exists(textBox1.Text))
                {
                    button4.Visible = false;
                }
                else
                {
                    button4.Visible = true;
                }
            }
            else
            {
                label1.Visible = false;
                textBox1.Visible = false;
                button3.Visible = false;
                button4.Visible = false;
            }

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                if (checkBox1.Checked)
                {
                    checkBox1.Checked = false;
                    label1.Visible = false;
                    textBox1.Visible = false;
                    button3.Visible = false;
                }

                textBox2.Visible = true;
                textBox3.Visible = true;
                button5.Visible = true;
                button6.Visible = true;
                label3.Visible = true;
                label4.Visible = true;

                if (textBox2.Text.Length == 0 || textBox2.Text.LastIndexOf(".") == -1 || textBox2.Text.LastIndexOf(".") < textBox2.Text.LastIndexOf("\\") || !File.Exists(textBox2.Text) || textBox3.Text.Length == 0 || textBox3.Text.LastIndexOf(".") == -1 || textBox3.Text.LastIndexOf(".") < textBox3.Text.LastIndexOf("\\") || !File.Exists(textBox3.Text))
                {
                    button4.Visible = false;
                }
                else
                {
                    button4.Visible = true;
                }

            }
            else
            {
                button4.Visible = false;
                textBox2.Visible = false;
                textBox3.Visible = false;
                button5.Visible = false;
                button6.Visible = false;
                label3.Visible = false;
                label4.Visible = false;
            }

        }


    }

    public class BPAHeader
    {
        public string cmpt;
        public string lines_qtdd;
        public string frames_qtdd;
        public string header_end;


        public BPAHeader(string ft_line)
        {
            cmpt = ft_line.Substring(7, 6);
            lines_qtdd = ft_line.Substring(13, 6);
            frames_qtdd = ft_line.Substring(19, 6);

            //Nesse espaço fica o Campo de controle


            //Nome do orgão de origem e etc...
            header_end = ft_line.Substring(29, ft_line.Length - 29);
        }

        public string getHeaderLine(string ctrl_child)
        {
            return "01#BPA#" + cmpt + lines_qtdd + frames_qtdd + ctrl_child + header_end;
        }
    }

    public class BPAC
    {
        public string cnes;
        public string cmpt;
        public string cbo;
        public string folha;
        public string seq;
        public string proc;
        public string idade;
        public string proc_qtdd;
        public string orgn;

        public BPAC(string raw_line)
        {
            cnes = raw_line.Substring(2, 7);
            cmpt = raw_line.Substring(9, 6);
            cbo = raw_line.Substring(15, 6);
            folha = raw_line.Substring(21, 3);
            seq = raw_line.Substring(24, 2);
            proc = raw_line.Substring(26, 10);
            idade = raw_line.Substring(36, 3);
            proc_qtdd = raw_line.Substring(39, 6);
            orgn = raw_line.Substring(45, 3);
        }

        public string getLine()
        {
            string drop = "02";
            drop += "\t";
            drop += cnes;
            drop += "\t";
            drop += cmpt;
            drop += "\t";
            drop += cbo;
            drop += "\t";
            drop += folha;
            drop += "\t";
            drop += seq;
            drop += "\t";
            drop += proc;
            drop += "\t";
            drop += idade;
            drop += "\t";
            drop += proc_qtdd;
            drop += "\t";
            drop += orgn;
            drop += "\n";

            return drop;
        }

        public ArrayList getArrayLine()
        {
            ArrayList drop = new ArrayList();
            drop.Add("02");
            drop.Add(cnes);
            drop.Add(cmpt);
            drop.Add(cbo);
            drop.Add(folha);
            drop.Add(seq);
            drop.Add(proc);
            drop.Add(idade);
            drop.Add(proc_qtdd);
            drop.Add(orgn);

            return drop;
        }
    }

    public class BPAI
    {
        public string cnes;
        public string cmpt;
        public string prof_cns;
        public string cbo;
        public string data;
        public string folha;
        public string seq;
        public string proc;
        public string pac_cns;
        public string sexo;
        public string ibge;
        public string cid;
        public string idade;
        public string qtdd;
        public string ca;      //Caracter de atendimento
        public string nae;     //Numero de autorização do estabelecimento
        public string orgn;        //Origem das informações
        public string pac_name;
        public string nsc_data;
        public string raca;
        public string etnia;
        public string nac;     //Nacionalidade
        public string srv_code;
        public string class_code;
        public string cse;         //Código da sequencia da equipe
        public string cae;         //Código da área da equipe
        public string cnpj;
        public string cep;
        public string log_type;
        public string log;           //Rua
        public string complemento;
        public string ln;            //n° de casa
        public string bairro;
        public string tel;
        public string email;
        public string ine;			//Identificação nacional da equipe


        public BPAI(string raw_line)
        {
            cnes = raw_line.Substring(2, 7);
            cmpt = raw_line.Substring(9, 6);
            prof_cns = raw_line.Substring(15, 15);
            cbo = raw_line.Substring(30, 6);
            data = raw_line.Substring(36, 8);
            folha = raw_line.Substring(44, 3);
            seq = raw_line.Substring(47, 2);
            proc = raw_line.Substring(49, 10);
            pac_cns = raw_line.Substring(59, 15);
            sexo = raw_line.Substring(74, 1);
            ibge = raw_line.Substring(75, 6);
            cid = raw_line.Substring(81, 4);
            idade = raw_line.Substring(85, 3);
            qtdd = raw_line.Substring(88, 6);
            ca = raw_line.Substring(94, 2);
            nae = raw_line.Substring(96, 13);
            orgn = raw_line.Substring(109, 3);
            pac_name = raw_line.Substring(112, 30);
            nsc_data = raw_line.Substring(142, 8);
            raca = raw_line.Substring(150, 2);
            etnia = raw_line.Substring(152, 4);
            nac = raw_line.Substring(156, 3);
            srv_code = raw_line.Substring(159, 3);
            class_code = raw_line.Substring(162, 3);
            cse = raw_line.Substring(165, 8);
            cae = raw_line.Substring(173, 4);
            cnpj = raw_line.Substring(177, 14);
            cep = raw_line.Substring(191, 8);
            log_type = raw_line.Substring(199, 3);
            log = raw_line.Substring(202, 30);
            complemento = raw_line.Substring(232, 10);
            ln = raw_line.Substring(242, 5);
            bairro = raw_line.Substring(247, 30);
            tel = raw_line.Substring(277, 11);
            email = raw_line.Substring(288, 40);
            ine = raw_line.Substring(328, 10);
        }

        public string getLine()
        {
            string drop = "03";
            drop += "\t";
            drop += cnes;
            drop += "\t";
            drop += cmpt;
            drop += "\t";
            drop += prof_cns;
            drop += "\t";
            drop += cbo;
            drop += "\t";
            drop += data;
            drop += "\t";
            drop += folha;
            drop += "\t";
            drop += seq;
            drop += "\t";
            drop += proc;
            drop += "\t";
            drop += pac_cns;
            drop += "\t";
            drop += sexo;
            drop += "\t";
            drop += ibge;
            drop += "\t";
            drop += cid;
            drop += "\t";
            drop += idade;
            drop += "\t";
            drop += qtdd;
            drop += "\t";
            drop += ca;
            drop += "\t";
            drop += nae;
            drop += "\t";
            drop += orgn;
            drop += "\t";
            drop += pac_name;
            drop += "\t";
            drop += nsc_data;
            drop += "\t";
            drop += raca;
            drop += "\t";
            drop += etnia;
            drop += "\t";
            drop += nac;
            drop += "\t";
            drop += srv_code;
            drop += "\t";
            drop += class_code;
            drop += "\t";
            drop += cse;
            drop += "\t";
            drop += cae;
            drop += "\t";
            drop += cnpj;
            drop += "\t";
            drop += cep;
            drop += "\t";
            drop += log_type;
            drop += "\t";
            drop += log;
            drop += "\t";
            drop += complemento;
            drop += "\t";
            drop += ln;
            drop += "\t";
            drop += bairro;
            drop += "\t";
            drop += tel;
            drop += "\t";
            drop += email;
            drop += "\t";
            drop += ine;
            drop += "\n";

            return drop;
        }

        public ArrayList getArrayLine()
        {
            ArrayList drop = new ArrayList();
            drop.Add("03");
            drop.Add(cnes);
            drop.Add(cmpt);
            drop.Add(prof_cns);
            drop.Add(cbo);
            drop.Add(data);
            drop.Add(folha);
            drop.Add(seq);
            drop.Add(proc);
            drop.Add(pac_cns);
            drop.Add(sexo);
            drop.Add(ibge);
            drop.Add(cid);
            drop.Add(idade);
            drop.Add(qtdd);
            drop.Add(ca);
            drop.Add(nae);
            drop.Add(orgn);
            drop.Add(pac_name);
            drop.Add(nsc_data);
            drop.Add(raca);
            drop.Add(etnia);
            drop.Add(nac);
            drop.Add(srv_code);
            drop.Add(class_code);
            drop.Add(cse);
            drop.Add(cae);
            drop.Add(cnpj);
            drop.Add(cep);
            drop.Add(log_type);
            drop.Add(log);
            drop.Add(complemento);
            drop.Add(ln);
            drop.Add(bairro);
            drop.Add(tel);
            drop.Add(email);
            drop.Add(ine);

            return drop;
        }
    }

    public class BPA
    {
        BPAHeader header;
        public ArrayList bpac = new ArrayList();
        public ArrayList bpai = new ArrayList();
        private bool isABPAFile(string filename)
        {
            StreamReader sr = new StreamReader(filename);
            string line = "" + sr.ReadLine();
            if (!line.StartsWith("01#"))
            {
                return false;
            }

            return true;
        }

        public BPA(string filename)
        {
            if (!isABPAFile(filename))
            {
                MessageBox.Show("Arquivo inválido", "Falha", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                //header = new BPAHeader(filename);

                StreamReader sr = new StreamReader(filename, System.Text.Encoding.Default);

                //Criando um backup do cabeçalho
                string line = sr.ReadLine();                //<<<Primeira linha do arquivo
                FileStream fs = File.Create("header.cfg");
                fs.Close();
                File.AppendAllText("header.cfg", line);
                //--------------------------------------------


                while ((line = sr.ReadLine()) != null)
                {
                    if (line.StartsWith("02"))
                    {
                        bpac.Add(new BPAC(line));
                    }
                    else if (line.StartsWith("03"))
                    {
                        bpai.Add(new BPAI(line));
                    }
                }

            }
        }

        public void dropBPAI2Excel(string filename)
        {
            FileStream fs = File.Create(filename);
            fs.Close();

            //Tipo de registro	CNES	Competencia	CNS do prof.	CBO	Data do atendimento	Folha	Sequencia	Procedimento	CNS do pac.	Sexo	IBGE	CID	Idade	Quantidade	Carac. do atend.	N. da aut. do estab.	Origem	Nome	Data de nasc.	Raça	Etnia	Nacionalidade	Cód. do serv.	Cód. da class.	Cód. da seq. da equipe	Cód. area da equipe	CNPJ	CEP	Tipo de logradouro	Endereço	Complemento	Casa	Bairro	Telefone	Email	INE
            File.AppendAllText(filename, "Tipo de registro\tCNES\tCompetencia\tCNS do prof.\tCBO\tData do atendimento\tFolha\tSequencia\tProcedimento\tCNS do pac.\tSexo\tIBGE\tCID\tIdade\tQuantidade\tCarac. do atend.\tN. da aut. do estab.\tOrigem\tNome\tData de nasc.\tRaça\tEtnia\tNacionalidade\tCód. do serv.\tCód. da class.\tCód. da seq. da equipe\tCód. area da equipe\tCNPJ\tCEP\tTipo de logradouro\tEndereço\tComplemento\tCasa\tBairro\tTelefone\tEmail\tINE\n");

            for (int i = 0; i < bpai.ToArray().Length; i++)
            {
                File.AppendAllText(filename, ((BPAI)bpai[i]).getLine());

            }

        }

        public void dropBPAC2Excel(string filename)
        {
            FileStream fs = File.Create(filename);
            fs.Close();

            //Tipo de registro	CNES	Competencia	CBO	Folha	Sequencia	Procedimento	Idade	Quantidade	Origem
            File.AppendAllText(filename, "Tipo de registro\tCNES\tCompetencia\tCBO\tFolha\tSequencia\tProcedimento\tIdade\tQuantidade\tOrigem\n");

            for (int i = 0; i < bpac.ToArray().Length; i++)
            {
                File.AppendAllText(filename, ((BPAC)bpac[i]).getLine());

            }

        }

        public void dropBPA2Excel(string bpac_filename, string bpai_filename, ProgressBar pb)
        {
            pb.Maximum = bpai.ToArray().Length + bpac.ToArray().Length;
            pb.Value = 0;

            FileStream fs = File.Create(bpac_filename);
            fs.Close();

            //Criando tabela
            XLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Sample Sheet");

            //Tipo de registro	CNES	Competencia	CBO	Folha	Sequencia	Procedimento	Idade	Quantidade	Origem
            File.AppendAllText(bpac_filename, "Tipo de registro\tCNES\tCompetencia\tCBO\tFolha\tSequencia\tProcedimento\tIdade\tQuantidade\tOrigem\n");

            ArrayList bpac_header = new ArrayList();
            bpac_header.Add("Tipo de registro");
            bpac_header.Add("CNES");
            bpac_header.Add("Competencia");
            bpac_header.Add("CBO");
            bpac_header.Add("Folha");
            bpac_header.Add("Sequencia");
            bpac_header.Add("Procedimento");
            bpac_header.Add("Idade");
            bpac_header.Add("Quantidade");
            bpac_header.Add("Origem");

            AddRow(worksheet, bpac_header, 0);

            for (int i = 0; i < bpac.ToArray().Length; i++)
            {
                File.AppendAllText(bpac_filename, ((BPAC)bpac[i]).getLine());                
                AddRow(worksheet, ((BPAC)bpac[i]).getArrayLine(), i + 1);
                pb.Value = pb.Value + 1;
            }

            workbook.SaveAs(bpac_filename.Replace(".txt", ".xlsx"));

            fs = File.Create(bpai_filename);
            fs.Close();

            //Criando tabela
            workbook = new XLWorkbook();
            worksheet = workbook.Worksheets.Add("Sample Sheet");

            //Tipo de registro	CNES	Competencia	CNS do prof.	CBO	Data do atendimento	Folha	Sequencia	Procedimento	CNS do pac.	Sexo	IBGE	CID	Idade	Quantidade	Carac. do atend.	N. da aut. do estab.	Origem	Nome	Data de nasc.	Raça	Etnia	Nacionalidade	Cód. do serv.	Cód. da class.	Cód. da seq. da equipe	Cód. area da equipe	CNPJ	CEP	Tipo de logradouro	Endereço	Complemento	Casa	Bairro	Telefone	Email	INE
            File.AppendAllText(bpai_filename, "Tipo de registro\tCNES\tCompetencia\tCNS do prof.\tCBO\tData do atendimento\tFolha\tSequencia\tProcedimento\tCNS do pac.\tSexo\tIBGE\tCID\tIdade\tQuantidade\tCarac. do atend.\tN. da aut. do estab.\tOrigem\tNome\tData de nasc.\tRaça\tEtnia\tNacionalidade\tCód. do serv.\tCód. da class.\tCód. da seq. da equipe\tCód. area da equipe\tCNPJ\tCEP\tTipo de logradouro\tEndereço\tComplemento\tCasa\tBairro\tTelefone\tEmail\tINE\n");


            ArrayList bpai_header = new ArrayList();
            bpai_header.Add("Tipo de registro");
            bpai_header.Add("CNES");
            bpai_header.Add("Competencia");
            bpai_header.Add("CNS do prof.");
            bpai_header.Add("CBO");
            bpai_header.Add("Data do atendimento");
            bpai_header.Add("Folha");
            bpai_header.Add("Sequencia");
            bpai_header.Add("Procedimento");
            bpai_header.Add("CNS do pac.");
            bpai_header.Add("Sexo");
            bpai_header.Add("IBGE");
            bpai_header.Add("CID");
            bpai_header.Add("Idade");
            bpai_header.Add("Quantidade");
            bpai_header.Add("Carac. do atend.");
            bpai_header.Add("N. da aut. do estab.");
            bpai_header.Add("Origem");
            bpai_header.Add("Nome");
            bpai_header.Add("Data de nasc.");
            bpai_header.Add("Raça");
            bpai_header.Add("Etnia");
            bpai_header.Add("Nacionalidade");
            bpai_header.Add("Cód. do serv.");
            bpai_header.Add("Cód. da class.");
            bpai_header.Add("Cód. da seq. da equipe");
            bpai_header.Add("Cód. area da equipe");
            bpai_header.Add("CNPJ");
            bpai_header.Add("CEP");
            bpai_header.Add("Tipo de logradouro");
            bpai_header.Add("Endereço");
            bpai_header.Add("Complemento");
            bpai_header.Add("Casa");
            bpai_header.Add("Bairro");
            bpai_header.Add("Telefone");
            bpai_header.Add("Email");
            bpai_header.Add("INE");

            AddRow(worksheet, bpai_header, 0);

            for (int i = 0; i < bpai.ToArray().Length; i++)
            {
                File.AppendAllText(bpai_filename, ((BPAI)bpai[i]).getLine());
                AddRow(worksheet, ((BPAI)bpai.ToArray()[i]).getArrayLine(), i + 1);
                pb.Value = pb.Value + 1;
            }

            workbook.SaveAs(bpai_filename.Replace(".txt", ".xlsx"));
        }

        private string GetNextCellName(string name)
        {
            string drop;

            if (name.Length == 1)
            {
                if (name.ToCharArray()[0] != 'Z')
                {
                    drop = "" + (char)((int)name.ToCharArray()[0] + 1);

                }
                else
                {
                    drop = "AA";
                }

                return drop;
            }
            else
            {
                if (name.ToCharArray()[name.Length - 1] != 'Z')
                {
                    char mod = (char)(name.ToCharArray()[name.Length - 1] + 1);
                    drop = name.Substring(0, name.Length - 1) + mod;
                }
                else
                {
                    drop = GetNextCellName(name.Substring(0, name.Length - 1)) + "A";
                }

                return drop;
            }

        }

        private void AddRow(IXLWorksheet worksheet, ArrayList row, int current_row)
        {
            string c = "A";
            string row_num = "" + (current_row + 1);
            string cell_name = "A";
            int rl = row.ToArray().Length;

            Console.Write("line 856: row.ToArray()[0] >> " + row.ToArray()[0]);

            for (int i = 0; i < rl; i++)
            {
                cell_name = c + row_num;
                worksheet.Cell(cell_name).Value = "'" + row.ToArray()[i];
                c = GetNextCellName(c);
            }

            

        }

    }

    public class BPABuilder
    {
        public ArrayList bpac = new ArrayList();
        public ArrayList bpai = new ArrayList();
        public BPAHeader header;
        public ulong fully_sum = 0;

        //Filename: Nome do arquivo de saída
        //bpac_filename: arquivo de entrada de registros BPAC
        //bpai_filename: arquivo de entrada de registros BPAI

        public BPABuilder(string filename, string bpac_filename, string bpai_filename)
        {
            StreamReader sr = new StreamReader("header.cfg");

            header = new BPAHeader(sr.ReadLine());

            XLWorkbook workbook = new XLWorkbook(bpac_filename);
            IXLWorksheet ws1 = workbook.Worksheet(1);

            string line;
            IXLRow row;
            //Percorrendo linhas da planilha com os registros do tipo BPAC
            for (int i = 1; i <= ws1.Rows().ToArray().Length; i++)
            {
                line = "";
                row = ws1.Row(i);

                //Percorrendo colunas da linha com o registro BPAC
                for (int j = 1; j <= row.Cells().ToArray().Length; j++)
                {
                    line += row.Cell(j).Value.ToString();
                }

                fully_sum += ulong.Parse(row.Cell(7).ToString()) + ulong.Parse(row.Cell(9).ToString());
                bpac.Add(line);
            }

            workbook = new XLWorkbook(bpai_filename);
            ws1 = workbook.Worksheet(1);

            //Percorrendo linhas da planilha com os registros do tipo BPAI
            for (int i = 1; i <= ws1.Rows().ToArray().Length; i++)
            {
                line = "";
                row = ws1.Row(i);

                //Percorrendo colunas da linha com o registro BPAC
                for (int j = 1; j <= row.Cells().ToArray().Length; j++)
                {
                    line += row.Cell(j).Value.ToString();
                }

                fully_sum += ulong.Parse(row.Cell(9).ToString()) + ulong.Parse(row.Cell(15).ToString());
                bpai.Add(line);
            }

            fully_sum = (fully_sum % 1111) + 1111;

            FileStream fs = File.Create(filename);
            fs.Close();

            File.AppendAllText(filename, header.getHeaderLine("" + fully_sum) + "\n");

            for (int i = 0; i < bpac.ToArray().Length; i++)
            {
                File.AppendAllText(filename, (string)bpac[i] + "\n");
            }

            for (int i = 0; i < bpai.ToArray().Length; i++)
            {
                File.AppendAllText(filename, (string)bpai[i] + "\n");
            }

        }

        public BPABuilder(string filename, string bpac_filename, string bpai_filename, ProgressBar pb)
        {
            Console.WriteLine("bp1");
            StreamReader sr = new StreamReader("header.cfg");

            header = new BPAHeader(sr.ReadLine());

            Console.WriteLine("bp2");

            XLWorkbook workbook = new XLWorkbook(bpac_filename);
            IXLWorksheet ws1 = workbook.Worksheet(1);

            Console.WriteLine("bp3");

            string line;
            IXLRow row;
            int ws_len = ws1.Rows().ToArray().Length;

            //Percorrendo linhas da planilha com os registros do tipo BPAC
            for (int i = 2; i <= ws_len; i++)
            {
                //Console.WriteLine("bp3.1");
                line = "";
                row = ws1.Row(i);

                //Console.WriteLine("bp3.2");

                Console.WriteLine("i = " + i + "/" + ws_len);

                //Console.WriteLine("bp3.3");
                //Percorrendo colunas da linha com o registro BPAC
                for (int j = 1; j <= row.Cells().ToArray().Length; j++)
                {
                    //Console.WriteLine("bp3.4");
                    //Console.WriteLine("j = " + j + "/" + row.Cells().ToArray().Length);
                    line += row.Cell(j).Value.ToString();
                    //Console.WriteLine("bp3.5");
                }

                //Console.WriteLine("bp3.6");

                fully_sum += ulong.Parse(row.Cell(7).Value.ToString()) + ulong.Parse(row.Cell(9).Value.ToString());

                Console.WriteLine("fully_sum = " + fully_sum);

                bpac.Add(line);

                //Console.WriteLine("bp3.8");
            }

            Console.WriteLine("bp4");
            workbook = new XLWorkbook(bpai_filename);
            ws1 = workbook.Worksheet(1);
            ws_len = ws1.Rows().ToArray().Length;

            //Percorrendo linhas da planilha com os registros do tipo BPAI
            for (int i = 2; i <= ws_len; i++)
            {
                line = "";
                row = ws1.Row(i);

                //Percorrendo colunas da linha com o registro BPAC
                for (int j = 1; j <= row.Cells().ToArray().Length; j++)
                {
                    line += row.Cell(j).Value.ToString();
                }

                fully_sum += ulong.Parse(row.Cell(9).Value.ToString()) + ulong.Parse(row.Cell(15).Value.ToString());

                Console.WriteLine("fully_sum = " + fully_sum);

                bpai.Add(line);
            }

            fully_sum = (fully_sum % 1111) + 1111;

            pb.Invoke((Action)delegate () {
                pb.Maximum = bpac.ToArray().Length + bpai.ToArray().Length;
                pb.Value = 0;
            });


            FileStream fs = File.Create(filename);
            fs.Close();

            File.AppendAllText(filename, header.getHeaderLine("" + fully_sum) + "\r\n");

            Console.WriteLine("bp5");

            for (int i = 0; i < bpac.ToArray().Length; i++)
            {
                try
                {
                    File.AppendAllText(filename, (string)bpac[i] + "\r\n");
                    pb.Invoke((Action)delegate () {
                        pb.Value = pb.Value + 1;
                    });
                }
                catch(IOException e)
                {
                    Console.Write("Line 1078 loop " + i + " Error: " + e.ToString());

                    --i;
                }

            }

            Console.WriteLine("bp6");

            for (int i = 0; i < bpai.ToArray().Length; i++)
            {
                try
                {
                    File.AppendAllText(filename, (string)bpai[i] + "\r\n");
                    pb.Invoke((Action)delegate () {
                        pb.Value = pb.Value + 1;
                    });
                }
                catch (IOException e)
                {
                    Console.Write("Line 1099 loop " + i + " Error: " + e.ToString());
                }
            }
        }

    }
}
