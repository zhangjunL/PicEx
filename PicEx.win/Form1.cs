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
            // 打开文件对话框
            OpenFileDialog openFileDialog = new OpenFileDialog();
            // 设置文件对话框的标题
            openFileDialog.Title = "请选择文件";
            // 设置文件对话框的初始目录
            openFileDialog.InitialDirectory = defaultOpenLocation;
            // 设置文件对话框的文件类型
            openFileDialog.Filter = "所有文件(*.*)|*.*|pdf文件(*.pdf)|*.pdf|word文件(*.docx)|*.docx|excel文件(*.xlsx)|*.xlsx|ppt文件(*.pptx)|*.pptx";
            // 设置文件对话框的默认文件类型
            openFileDialog.FilterIndex = 0;
            // 设置对话框是否记忆之前打开的目录
            openFileDialog.RestoreDirectory = true;
            // 如果用户点击了确定按钮，则打开文件
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // 获取用户选择的文件，并判断文件是否存在
                defaultOpenLocation = openFileDialog.FileName;
                
                //string filePath = openFileDialog.FileName;
                textBox1.Text = defaultOpenLocation;
                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 打开文件夹对话框
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            // 设置文件夹对话框的描述信息
            folderBrowserDialog.Description = "请选择文件夹";
            // 设置文件夹对话框的初始目录
            folderBrowserDialog.SelectedPath = defaultSaveLocation;
            // 设置是否显示新建文件夹按钮
            folderBrowserDialog.ShowNewFolderButton = true;
            // 如果用户点击了确定按钮，则打开文件夹
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                // 获取用户选择的文件夹，并判断文件夹是否存在
                defaultSaveLocation = folderBrowserDialog.SelectedPath;
                textBox2.Text = defaultSaveLocation;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filePath = defaultOpenLocation;
            string outputDirectory = defaultSaveLocation;
            //如果filePath或 outputDirectory 为空，则抛出异常
            if (string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(outputDirectory))
            {
                MessageBox.Show("请选择文件和保存路径");
                return;
            }
            // 如果文件存在，则进行解析
            if (File.Exists(filePath))
            {
                // 获取文件扩展名
                string extension = Path.GetExtension(filePath);
                // 如果扩展名不是pdf，则抛出异常
                switch (extension)
                {
                    case ".pdf":
                        PdfImageExtractor.ExtractImagesFromPdf(filePath, outputDirectory);
                        MessageBox.Show("提取成功");
                        break;
                    case ".docx":
                        ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputDirectory,"word");
                        MessageBox.Show("提取成功");
                        break;
                    case ".xlsx":
                        ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputDirectory,"xl");
                        MessageBox.Show("提取成功");
                        break;
                    case ".pptx":
                        ExcelImageExtractor.ExtractImagesFromExcel(filePath, outputDirectory, "ppt");
                        MessageBox.Show("提取成功");
                        break;
                    default:
                        MessageBox.Show("该文件类型暂不支持解析");
                        break;
                        //throw new ArgumentException("该文件类型暂不支持解析");

                }
            }
            else
            {
                MessageBox.Show("文件不存在");
            }
        }
    }
}