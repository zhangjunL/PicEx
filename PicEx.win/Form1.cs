using PicEx.Office;
using PicEx.Pdf;


namespace PicEx.win
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox2.Text = defaultSaveLocation;
            textBox1.Text = defaultOpenLocation;
            
        }
        private string defaultSaveLocation = @"";

        private string defaultOpenLocation = @"";


       
        private void button1_Click(object sender, EventArgs e)
        {
            // ���ļ��Ի���
            OpenFileDialog openFileDialog = new OpenFileDialog();
            // �����ļ��Ի���ı���
            openFileDialog.Title = "��ѡ���ļ�";
            // �����ļ��Ի���ĳ�ʼĿ¼
            openFileDialog.InitialDirectory = defaultOpenLocation;
            // �����ļ��Ի�����ļ�����
            openFileDialog.Filter = "�����ļ�(*.*)|*.*|pdf�ļ�(*.pdf)|*.pdf|word�ļ�(*.docx)|*.docx|excel�ļ�(*.xlsx)|*.xlsx|ppt�ļ�(*.pptx)|*.pptx";
            // �����ļ��Ի����Ĭ���ļ�����
            openFileDialog.FilterIndex = 0;
            // ���öԻ����Ƿ����֮ǰ�򿪵�Ŀ¼
            openFileDialog.RestoreDirectory = true;
            // ����û������ȷ����ť������ļ�
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // ��ȡ�û�ѡ����ļ������ж��ļ��Ƿ����
                defaultOpenLocation = openFileDialog.FileName;
                
                //string filePath = openFileDialog.FileName;
                textBox1.Text = defaultOpenLocation;
                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // ���ļ��жԻ���
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            // �����ļ��жԻ����������Ϣ
            folderBrowserDialog.Description = "��ѡ���ļ���";
            // �����ļ��жԻ���ĳ�ʼĿ¼
            folderBrowserDialog.SelectedPath = defaultSaveLocation;
            // �����Ƿ���ʾ�½��ļ��а�ť
            folderBrowserDialog.ShowNewFolderButton = true;
            // ����û������ȷ����ť������ļ���
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                // ��ȡ�û�ѡ����ļ��У����ж��ļ����Ƿ����
                defaultSaveLocation = folderBrowserDialog.SelectedPath;
                textBox2.Text = defaultSaveLocation;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filePath = defaultOpenLocation;
            string outputDirectory = defaultSaveLocation;
            //���filePath�� outputDirectory Ϊ�գ����׳��쳣
            if (string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(outputDirectory))
            {
                MessageBox.Show("��ѡ���ļ��ͱ���·��");
                return;
            }
            // ����ļ����ڣ�����н���
            if (File.Exists(filePath))
            {
                // ��ȡ�ļ���չ��
                string extension = Path.GetExtension(filePath);
                // �����չ������pdf�����׳��쳣
                switch (extension)
                {
                    case ".pdf":
                        PdfImageExtractor.ExtractImagesFromPdf(filePath, outputDirectory);
                        MessageBox.Show("��ȡ�ɹ�");
                        break;
                    case ".docx":
                        ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputDirectory,"word");
                        MessageBox.Show("��ȡ�ɹ�");
                        break;
                    case ".xlsx":
                        ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputDirectory,"xl");
                        MessageBox.Show("��ȡ�ɹ�");
                        break;
                    case ".pptx":
                        ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputDirectory, "ppt");
                        MessageBox.Show("��ȡ�ɹ�");
                        break;
                    default:
                        MessageBox.Show("���ļ������ݲ�֧�ֽ���");
                        break;
                        //throw new ArgumentException("���ļ������ݲ�֧�ֽ���");

                }
            }
            else
            {
                MessageBox.Show("�ļ�������");
            }
        }
    }
}