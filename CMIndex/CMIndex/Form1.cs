using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Threading;

namespace CMIndex
{
    public partial class Form1 : Form
    {
        bool isVoice = true;
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: 这行代码将数据加载到表“herbalDataSet1.中药”中。您可以根据需要移动或删除它。
            this.中药TableAdapter.Fill(this.herbalDataSet1.中药);
        }

        private void toolStripTextBox1_TextChanged(object sender, EventArgs e)
        {
            toolStripComboBox1.DroppedDown = false;
            toolStripComboBox1.Items.Clear();
            toolStripComboBox1.Text = "";

            if (IsEnCh(toolStripTextBox1.Text))
            {
                IndexCMNameByPY(toolStripTextBox1.Text);
                toolStripComboBox1.DroppedDown = true;

                this.Cursor = Cursors.Default;
            }
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            IndexEachOtherData(toolStripComboBox1.Text);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (isVoice == true)
            {
                toolStripButton1.Image = Image.FromFile(Environment.CurrentDirectory + "/静音.png");
                isVoice = false;
            }
            else
            {
                toolStripButton1.Image = Image.FromFile(Environment.CurrentDirectory + "/声音.png");
                isVoice = true;
            }
        }

        //检索
        private void IndexCMNameByPY(string str)
        {
            str = str.ToUpper();
            int len = str.Length;
            List<string> strs = new List<string>();

            switch (len)
            {
                case 1:
                    for (int j = 0; j < dataGridView1.RowCount; j++)
                    {
                        if (dataGridView1[2, j].Value != null && dataGridView1[2, j].Value.ToString() != "")
                        {
                            if (str[0] == dataGridView1[2, j].Value.ToString()[0])
                            {
                                strs.Add(dataGridView1[1, j].Value.ToString());
                            }
                        }
                    }
                    //去重
                    HashSet<string> hs = new HashSet<string>(strs);  //此时已经去掉重复的数据保存在hashset中
                    toolStripComboBox1.Items.AddRange((new List<string>(hs)).ToArray());
                    break;
                case 2:
                    for (int j = 0; j < dataGridView1.RowCount; j++)
                    {
                        if (dataGridView1[2, j].Value != null && dataGridView1[2, j].Value.ToString() != "" && dataGridView1[2, j].Value.ToString().Length >= 2)
                        {
                            if (str[0] == dataGridView1[2, j].Value.ToString()[0] && str[1] == dataGridView1[2, j].Value.ToString()[1])
                            {
                                strs.Add(dataGridView1[1, j].Value.ToString());
                            }
                        }
                    }
                    //去重
                    HashSet<string> hs1 = new HashSet<string>(strs);  //此时已经去掉重复的数据保存在hashset中
                    toolStripComboBox1.Items.AddRange((new List<string>(hs1)).ToArray());
                    break;
                case 3:
                    for (int j = 0; j < dataGridView1.RowCount; j++)
                    {
                        if (dataGridView1[2, j].Value != null && dataGridView1[2, j].Value.ToString() != "" && dataGridView1[2, j].Value.ToString().Length >= 3)
                        {
                            if (str[0] == dataGridView1[2, j].Value.ToString()[0] && str[1] == dataGridView1[2, j].Value.ToString()[1] && str[2] == dataGridView1[2, j].Value.ToString()[2])
                            {
                                strs.Add(dataGridView1[1, j].Value.ToString());
                            }
                        }
                    }
                    //去重
                    HashSet<string> hs2 = new HashSet<string>(strs);  //此时已经去掉重复的数据保存在hashset中
                    toolStripComboBox1.Items.AddRange((new List<string>(hs2)).ToArray());
                    break;
                case 4:
                    for (int j = 0; j < dataGridView1.RowCount; j++)
                    {
                        if (dataGridView1[2, j].Value != null && dataGridView1[2, j].Value.ToString() != "" && dataGridView1[2, j].Value.ToString().Length >= 4)
                        {
                            if (str[0] == dataGridView1[2, j].Value.ToString()[0] && str[1] == dataGridView1[2, j].Value.ToString()[1] && str[2] == dataGridView1[2, j].Value.ToString()[2] && str[3] == dataGridView1[2, j].Value.ToString()[3])
                            {
                                strs.Add(dataGridView1[1, j].Value.ToString());
                            }
                        }
                    }
                    //去重
                    HashSet<string> hs3 = new HashSet<string>(strs);  //此时已经去掉重复的数据保存在hashset中
                    toolStripComboBox1.Items.AddRange((new List<string>(hs3)).ToArray());
                    break;
                case 5:
                    for (int j = 0; j < dataGridView1.RowCount; j++)
                    {
                        if (dataGridView1[2, j].Value != null && dataGridView1[2, j].Value.ToString() != "" && dataGridView1[2, j].Value.ToString().Length >= 5)
                        {
                            if (str[0] == dataGridView1[2, j].Value.ToString()[0] && str[1] == dataGridView1[2, j].Value.ToString()[1] && str[2] == dataGridView1[2, j].Value.ToString()[2] && str[3] == dataGridView1[2, j].Value.ToString()[3] && str[4] == dataGridView1[2, j].Value.ToString()[4])
                            {
                                strs.Add(dataGridView1[1, j].Value.ToString());
                            }
                        }
                    }
                    //去重
                    HashSet<string> hs4 = new HashSet<string>(strs);  //此时已经去掉重复的数据保存在hashset中
                    toolStripComboBox1.Items.AddRange((new List<string>(hs4)).ToArray());
                    break;
                default:
                    break;
            }
        }

        //检索对应数据源
        private void IndexEachOtherData(string str)
        {
            ClearColor();
            dataGridView1.Refresh();
            dataGridView1.ClearSelection();

            List<string> CMstrs = new List<string>();

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1[1, i].Value != null && dataGridView1[1, i].Value.ToString() != "")
                {
                    if (dataGridView1[1, i].Value.ToString() == str)
                    {
                        dataGridView1.Rows[i].Selected = true;
                        dataGridView1.CurrentCell = dataGridView1.Rows[i].Cells[1];
                        dataGridView1.CurrentRow.Selected = true;

                        int index = dataGridView1.CurrentRow.Index;
                        string ID = dataGridView1.Rows[index].Cells[0].Value.ToString();
                        string NAME = dataGridView1.Rows[index].Cells[1].Value.ToString(); ;

                        if (dataGridView1.Rows[index].Cells[4].Value != null && dataGridView1.Rows[index].Cells[4].Value.ToString() != "")
                        {
                            NAME = dataGridView1.Rows[index].Cells[4].Value.ToString();
                        }

                        CMLocation(ID);

                        if (isVoice)
                        {
                            string CMstr = NAME + "在" + ID[0] + "柜" + ID[1] + "杭" + ID[2] + "列";
                            CMstrs.Add(CMstr);
                        }
                    }
                }
            }

            if (isVoice)
            {
                string CMinWhere = "";

                for (int i = 0; i < CMstrs.ToArray().Length; i++)
                {
                    CMinWhere += CMstrs.ToArray()[i] + "。";
                }

                Thread t = new Thread(new ParameterizedThreadStart(PlayVoiceFromString));
                t.IsBackground = true;
                t.Start(CMinWhere);
            } 
        }

        //判断是否为字母
        private bool IsEnCh(string input)
        {
            string pattern = @"^[A-Za-z]+$";
            Regex regex = new Regex(pattern);

            return regex.IsMatch(input);
        }

        //位置标记
        private void CMLocation(String ID)
        {
            string GID = new string(new char[] { 'G', ID[0] });
            string YID = new string(new char[] { ID[3], ID[4] });
            Color color = ProductRandomColor(int.Parse(YID));

            foreach (Control ctrl in splitContainer2.Panel1.Controls)
            {
                if (ctrl is Button)
                {
                    if (((Button)ctrl).Text == GID)
                    {
                        ((Button)ctrl).BackColor = color;
                    } 
                }
            }

            foreach (Control ctrl in splitContainer2.Panel2.Controls)
            {
                if (ctrl is Button)
                {
                    if (((Button)ctrl).Text == YID)
                    {
                        ((Button)ctrl).BackColor = color;
                    }
                }
            }
        }

        //产生随机颜色
        private Color ProductRandomColor(int seed)
        {
            Random ro = new Random(seed);
            long tick = DateTime.Now.Ticks;
            Random ran = new Random((int)(tick & 0xffffffffL) | (int)(tick >> 32));

            int R = ran.Next(255);
            int G = ran.Next(255);
            int B = ran.Next(255);
            B = (R + G > 400) ? R + G - 400 : B;  //0 : 380 - R - G;
            B = (B > 255) ? (seed * 5) : B;

            return Color.FromArgb(R, G, B);
        }

        //颜色清除
        private void ClearColor()
        {
            foreach (Control ctrl in splitContainer2.Panel1.Controls)
            {
                if (ctrl is Button)
                {
                    ((Button)ctrl).BackColor = Color.White;
                }
            }

            foreach (Control ctrl in splitContainer2.Panel2.Controls)
            {
                if (ctrl is Button)
                {
                    ((Button)ctrl).BackColor = Color.White;
                }
            }
        }

        //语音播放
        private void PlayVoiceFromString(object ostr)
        {           
            string str = ostr.ToString();

            SpeechSynthesizer speech = new SpeechSynthesizer();
            speech.Rate = 0;  //语速为正常语速
            speech.Volume = 100;  //音量为最大音量

            speech.Speak(str);  //语音阅读方法
        }
    }
}
