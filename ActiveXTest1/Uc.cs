using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing.Printing;
using System.IO;
using System.Drawing;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace ActiveXTest1
{
    [Guid("073A987E-2A7C-4874-8BEE-321E04F4E84E")]
    public partial class Uc : UserControl, IObjectSafety
    {
        PrintDocument printDocument;
        PrintPreviewDialog printPreview;
        StringReader lineReader = null;
        private int foodLength;
        private int countLength;
        private int moneyLength;
        Ticket deserializedTicket;
        Boolean isKitchen;
        public Uc()
        {
            InitializeComponent();
            printDocument = new PrintDocument();
            printPreview = new PrintPreviewDialog();
            Margins margin = new Margins(1, 1, 1, 1);
            printDocument.DefaultPageSettings.Margins = margin;
            printDocument.DefaultPageSettings.PaperSize = new PaperSize("Custum", getInch(100), getInch(110));
            printDocument.PrintPage += new PrintPageEventHandler(this.printDocument_PrintPage);
            foodLength = 10;
            countLength = 7;
            moneyLength = 8;
        }

        #region IObjectSafety 成员
        private const string _IID_IDispatch = "{00020400-0000-0000-C000-000000000046}";
        private const string _IID_IDispatchEx = "{a6ef9860-c720-11d0-9337-00a0c90dcaa9}";
        private const string _IID_IPersistStorage = "{0000010A-0000-0000-C000-000000000046}";
        private const string _IID_IPersistStream = "{00000109-0000-0000-C000-000000000046}";
        private const string _IID_IPersistPropertyBag = "{37D84F60-42CB-11CE-8135-00AA004BB851}";

        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x00000001;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x00000002;
        private const int S_OK = 0;
        private const int E_FAIL = unchecked((int)0x80004005);
        private const int E_NOINTERFACE = unchecked((int)0x80004002);

        private bool _fSafeForScripting = true;
        private bool _fSafeForInitializing = true;

        public int GetInterfaceSafetyOptions(ref Guid riid, ref int pdwSupportedOptions, ref int pdwEnabledOptions)
        {
            int Rslt = E_FAIL;

            string strGUID = riid.ToString("B");
            pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            switch (strGUID)
            {
                case _IID_IDispatch:
                case _IID_IDispatchEx:
                    Rslt = S_OK;
                    pdwEnabledOptions = 0;
                    if (_fSafeForScripting == true)
                        pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER;
                    break;
                case _IID_IPersistStorage:
                case _IID_IPersistStream:
                case _IID_IPersistPropertyBag:
                    Rslt = S_OK;
                    pdwEnabledOptions = 0;
                    if (_fSafeForInitializing == true)
                        pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_DATA;
                    break;
                default:
                    Rslt = E_NOINTERFACE;
                    break;
            }

            return Rslt;
        }

        public int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions)
        {
            int Rslt = E_FAIL;
            string strGUID = riid.ToString("B");
            switch (strGUID)
            {
                case _IID_IDispatch:
                case _IID_IDispatchEx:
                    if (((dwEnabledOptions & dwOptionSetMask) == INTERFACESAFE_FOR_UNTRUSTED_CALLER) && (_fSafeForScripting == true))
                        Rslt = S_OK;
                    break;
                case _IID_IPersistStorage:
                case _IID_IPersistStream:
                case _IID_IPersistPropertyBag:
                    if (((dwEnabledOptions & dwOptionSetMask) == INTERFACESAFE_FOR_UNTRUSTED_DATA) && (_fSafeForInitializing == true))
                        Rslt = S_OK;
                    break;
                default:
                    Rslt = E_NOINTERFACE;
                    break;
            }

            return Rslt;
        }

        #endregion

        public void getData(String json)
        {
            deserializedTicket = JsonConvert.DeserializeObject<Ticket>(json);
            isKitchen = deserializedTicket.isKitchen;
            Boolean isPreview = deserializedTicket.isPreview;
            String printName = deserializedTicket.printName;
            printDocument.PrinterSettings.PrinterName = printName;
            printPreview.Document = printDocument;
            lineReader = new StringReader(isKitchen ? getKitchenPrintString(deserializedTicket) : getTicketPrintString(deserializedTicket));

            try
            {
                if (isPreview)
                {
                    if (printPreview.ShowDialog() == DialogResult.OK)
                    {
                        printDocument.Print();
                    }
                }
                else
                {
                    printDocument.Print();
                }
            }
            catch (Exception excep)
            {
                MessageBox.Show(excep.Message, "打印出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
                printDocument.PrintController.OnEndPrint(printDocument, new PrintEventArgs());
            }
        }

        private void printDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            Graphics g = e.Graphics; //获得绘图对象
            float yPosition = 0;   //绘制字符串的纵向位置
            int count = 0; //行计数器
            float leftMargin = e.MarginBounds.Left; //左边距
            float topMargin = e.MarginBounds.Top; //上边距
            string line = null; //行字符串
            Font printFont = isKitchen ? new Font(new FontFamily("宋体"), 14) : new Font(new FontFamily("宋体"), 10);
            Font notesFont = new Font(new FontFamily("宋体"), 12);
            Font typeFont = new Font(new FontFamily("黑体"), 20);
            SolidBrush brush = new SolidBrush(Color.Black); //刷子
            //逐行的循环打印
            while ((line = lineReader.ReadLine()) != null)
            {
                yPosition = topMargin + (count * printFont.GetHeight(g));
                if (line.StartsWith("注意"))
                {
                    g.DrawString(line, notesFont, brush, leftMargin, yPosition, new StringFormat());
                }
                else if (line.StartsWith("点菜") || line.StartsWith("加菜") || line.StartsWith("退菜"))
                {
                    g.DrawString(line, typeFont, brush, leftMargin, yPosition, new StringFormat());
                }
                else
                {
                    g.DrawString(line, printFont, brush, leftMargin, yPosition, new StringFormat());
                }
                count++;
            }
            lineReader = new StringReader(isKitchen ? getKitchenPrintString(deserializedTicket) : getTicketPrintString(deserializedTicket));
        }

        private int getInch(double cm)
        {
            return (int)(cm / 25.4) * 100;
        }

        public string getKitchenPrintString(Ticket kitchen)
        {
            StringBuilder sb = new StringBuilder();
            // 桌号
            string desk = kitchen.desk;
            string type = kitchen.type;
            List<Ticket.Menu> menu = kitchen.menu;
            int count = menu.Count;
            String notes;
            sb.Append("       " + desk + "号桌菜单" + "\n");
            sb.Append(type + "\n");
            sb.Append("---------------------------------------------------------------\n");
            sb.Append("时间 " + DateTime.Now.ToLongTimeString() + "\n");
            sb.Append("---------------------------------------------------------------\n");
            sb.Append(padRightTrueLen(" 菜名", foodLength, ' ') + padRightTrueLen("桌号", countLength, ' ') +
                padRightTrueLen("数量", countLength, ' ') + "\n");
            sb.Append("\n");
            for (int i = 0; i < menu.Count; i++)
            {
                sb.Append(padRightTrueLen(menu[i].name, foodLength, ' ') + padRightTrueLen(" " + desk, countLength, ' ') +
                    padRightTrueLen(" " + menu[i].count, countLength, ' ') + "\n");
                notes = menu[i].notes;
                if (notes != null)
                {
                    sb.Append(menu[i].notes + "\n");
                }
                sb.Append("\n");
            }
            sb.Append("---------------------------------------------------------------\n");
            sb.Append("数量: " + count + "\n");
            return sb.ToString();
        }

        public string getTicketPrintString(Ticket ticket)
        {
            StringBuilder sb = new StringBuilder();
            string restaurant = ticket.restaurant;
            string orderNo = ticket.orderNo;
            string address = ticket.address;
            string pay = ticket.pay;
            string telephone = ticket.telephone;
            string mobilephone = ticket.mobilephone;
            List<Ticket.Menu> menu = ticket.menu;
            int count = menu.Count;
            decimal cost = 0.00M;
            sb.Append("         " + restaurant + "\n");
            sb.Append("\n");
            sb.Append("---------------------------------------------------------------\n");
            sb.Append("日期:" + DateTime.Now.ToShortDateString() + "  " + "单号:" + orderNo + "\n");
            sb.Append("---------------------------------------------------------------\n");
            sb.Append(padRightTrueLen(" 菜名", foodLength, ' ') + padRightTrueLen("数量", countLength, ' ') +
                padRightTrueLen("单价", moneyLength, ' ') + "小计" + "\n");
            for (int i = 0; i < count; i++)
            {
                decimal sum = Decimal.Parse(menu[i].count) * Decimal.Parse(menu[i].price);
                sb.Append(padRightTrueLen((menu[i].name), foodLength, ' ') +
                    padRightTrueLen((" " + menu[i].count), countLength, ' ') +
                    padRightTrueLen((menu[i].price), moneyLength, ' ') + sum + "\n");
                cost += sum;
            }
            sb.Append("---------------------------------------------------------------\n");
            sb.Append("数量:" + count + "   合计:" + cost + "\n");
            sb.Append("付款: 现金" + " " + pay);
            sb.Append("   现金找零:" + " " + (Decimal.Parse(pay) - cost) + "\n");
            sb.Append("---------------------------------------------------------------\n");
            sb.Append("地址：" + address + "\n");
            sb.Append("电话：" + telephone + "   手机：" + mobilephone + "\n");
            sb.Append("\n");
            sb.Append("         谢谢惠顾，欢迎下次光临         ");
            sb.Append("\n");
            return sb.ToString();
        }

        // 根据asc码来判断字符串的长度，在0~127间字符长度加1，否则加2
        // 需要返回长度的字符串
        public int trueLength(string str)
        {
            int lenTotal = 0;
            int n = str.Length;
            string strWord = "";  //清空字符串
            int asc;
            for (int i = 0; i < n; i++)
            {
                strWord = str.Substring(i, 1);
                asc = Convert.ToChar(strWord);
                if (asc < 0 || asc > 127)      // 在0~127间字符长度加1，否则加2
                {
                    lenTotal = lenTotal + 2;
                }
                else
                {
                    lenTotal = lenTotal + 1;
                }
            }
            return lenTotal;
        }

        // 统一字符串的长度
        // 初始字符串
        // 规定统一字符串的长度
        // 追加的字符为' '
        // 返回统一后的字符串
        public string padRightTrueLen(string strOriginal, int maxTrueLength, char chrPad)
        {
            string strNew = strOriginal;
            if (strOriginal == null || maxTrueLength <= 0)
            {
                strNew = "";
                return strNew;
            }
            int trueLen = trueLength(strOriginal);
            if (trueLen > maxTrueLength)  // 如果字符串大于规定长度 将规定长度等于字符串长度
            {
                for (int i = 0; i < trueLen - maxTrueLength; i++)
                {
                    maxTrueLength += chrPad.ToString().Length;
                }
            }
            else //填充小于规定长度 用' '追加，直至等于规定长度
            {
                for (int i = 0; i < maxTrueLength - trueLen; i++)
                {
                    strNew += chrPad.ToString();
                }
            }
            return strNew;
        }

    }

}