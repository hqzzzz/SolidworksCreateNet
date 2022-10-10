using Newtonsoft.Json;
using PSWStandalone;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

namespace SolidNet
{
    public partial class CreateNet : Form
    {


        public static ISldWorks SwApp { get; private set; }
        static List<string> LineNames = new System.Collections.Generic.List<string>();
        SwNetData swNetData = new SwNetData();

        DataTable dt1 = new DataTable();
        DataTable dt2 = new DataTable();

        public CreateNet()
        {

            InitializeComponent();

            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is System.Windows.Forms.TextBox)
                {
                   ctrl.MouseCaptureChanged+= new EventHandler(WTextChanged);
                }

            }



            DataColumn dt1_col1 = new DataColumn("间距",typeof(double));
            DataColumn dt1_col2 = new DataColumn("数量", typeof(int));

      

            DataColumn dt2_col1 = new DataColumn("间距", typeof(double));
            DataColumn dt2_col2 = new DataColumn("数量", typeof(int));

            dt1.Columns.Add(dt1_col1);
            dt1.Columns.Add(dt1_col2);
            DataRow row1 = dt1.NewRow();
            dt1.Rows.Add(row1);


            dt2.Columns.Add(dt2_col1);
            dt2.Columns.Add(dt2_col2);
            DataRow row21 = dt2.NewRow();

            dt2.Rows.Add(row21);
    

            dataGrid_View1.DataSource= dt1;
            dataGrid_View1.Rows[0].Cells[0].Value = "60";
            dataGrid_View1.Rows[0].Cells[1].Value = "16";
            for (int i = 0; i < this.dataGrid_View1.Columns.Count; i++)
            {
                this.dataGrid_View1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            dataGrid_View1.AutoGenerateColumns = true;
            dataGrid_View1.RowHeadersVisible = false;
            dataGrid_View1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGrid_View1.AllowUserToResizeColumns = false;
            dataGrid_View1.AllowUserToResizeRows = false;



           
            dataGrid_View2.DataSource = dt2;
            for (int i = 0; i < this.dataGrid_View2.Columns.Count; i++)
            {
                this.dataGrid_View2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            dataGrid_View2.Rows[0].Cells[0].Value = "50.0";
            dataGrid_View2.Rows[0].Cells[1].Value = "10";
            dataGrid_View2.AutoGenerateColumns = true;
            dataGrid_View2.RowHeadersVisible = false;
            dataGrid_View2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGrid_View2.AllowUserToResizeColumns = false;
            dataGrid_View2.AllowUserToResizeRows = false;


            swNetData.D1 = 3.0;
            tB_D1.DataBindings.Add("Text", swNetData, "D1", true, DataSourceUpdateMode.OnPropertyChanged,null,"0.0");

            swNetData.D2 = 3.0;
            tB_D2.DataBindings.Add("Text", swNetData, "D2", true, DataSourceUpdateMode.OnPropertyChanged,null, "0.0");

            //tB_D2.Text = "3.0";
            //tB_D1.Text = "3.0";

           

            swNetData.H = 500;
            tB_H.DataBindings.Add("Text", swNetData, "H", false, DataSourceUpdateMode.OnPropertyChanged);
            swNetData.Hs1 = 30;
            tB_Hs1.DataBindings.Add("Text", swNetData, "Hs1", false, DataSourceUpdateMode.OnPropertyChanged);
            swNetData.Hs2 = 0;
            tB_Hs2.DataBindings.Add("Text", swNetData, "Hs2", false, DataSourceUpdateMode.OnPropertyChanged);
            swNetData.Hd1 = 90;
            tB_Hd1.DataBindings.Add("Text", swNetData, "Hd1", false, DataSourceUpdateMode.OnPropertyChanged);
            swNetData.Hd2 = 120;
            tB_Hd2.DataBindings.Add("Text", swNetData, "Hd2", false, DataSourceUpdateMode.OnPropertyChanged);
            swNetData.L1= 964;
            tB_L1.DataBindings.Add("Text", swNetData, "L1", false, DataSourceUpdateMode.OnPropertyChanged);

        }

        public static string str_LinName(int id)
        {

            return LineNames[id].ToString();
        }

        public static void add_LinName()
        {
            int tco =LineNames.Count+1;
            LineNames.Add($"Line{tco.ToString()}");
        
        }




        public static SldWorks ConnectToSolidWorks()
        {
            return PStandAlone.GetSolidWorks();
        }



        Dimension tp_dim = null;
        private void BT_TS_Click(object sender, EventArgs e)
        {

            //连接到Solidworks
            SldWorks swApp = CreateNet.ConnectToSolidWorks();
            Random random=new Random();
            LineNames.Clear();


 

            //存储先前设置
            bool InputDimFlag=swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);

            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            if (swApp != null)
            {
                //创建模型
                string partDefaultTemplate = swApp.GetDocumentTemplate((int)swDocumentTypes_e.swDocPART, random.Next().ToString(), 0, 0, 0);
                var newDoc = swApp.NewDocument(partDefaultTemplate, 0, 0, 0);
       

                if (newDoc != null)
                {
                    //链接到模型
                    ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;

                    //创建草图平面
                    bool boolstatus = swModel.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, false, 0, null, 0);
                    swModel.SketchManager.InsertSketch(true);

                    //绘制直线
                    var skSegment = swModel.SketchManager.CreateLine(0, 0, 0, swNetData.h, 0, 0);
                    add_LinName();

                    swModel.Extension.SelectByID2(str_LinName(0), "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                    swModel.AddDimension2(5E-02, -0.02, 5E-02);

                    //设置尺寸
                    Dimension dimension = (Dimension)swModel.Parameter("D1@草图1");
                    dimension.SystemValue = swNetData.h; 
                  
                  

                    if(swNetData.hs1 != 0)
                    {

                        double Lsin = Math.Abs(swNetData.hs1 / Math.Sin(SMath.toRadians(swNetData.Hd1)));
                        double Lx = Lsin * Math.Cos(SMath.toRadians(swNetData.Hd1));
                        double Ly = Lsin * Math.Sin(SMath.toRadians(swNetData.Hd1));
                        swModel.SketchManager.CreateLine(0, 0, 0, Lx, Ly, 0);
                        add_LinName();
                        swModel.Extension.SelectByID2(str_LinName(LineNames.Count - 1), "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                        swModel.Extension.SelectByID2(str_LinName(0), "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
                        swModel.AddDimension2(Lx - (2E-02), 0, 1E-02);
                        tp_dim = (Dimension)swModel.Parameter("D2@草图1");
                        tp_dim.Name = "HDelta1";
                        tp_dim.SystemValue = SMath.toRadians(swNetData.hd1);


                        swModel.Extension.SelectByID2(str_LinName(LineNames.Count-1), "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                        if (swModel.AddVerticalDimension2(-2E-03, 0, -1E-03)!=null)
                        {
                            tp_dim = (Dimension)swModel.Parameter("D2@草图1");
                            tp_dim.Name = "Hs1";
                            tp_dim.SystemValue = swNetData.hs1;
                        }

          
  
                        swModel.ClearSelection2(true);
                        swModel.Extension.SelectByID2("Point1", "SKETCHPOINT", 0, 0, 0, false, 0, null, 0);
                        swModel.SketchManager.CreateFillet(swNetData.d1, (int)swConstrainedCornerAction_e.swConstrainedCornerDeleteGeometry);

                    }

                    if (swNetData.hs2!=0)
                    {
                        double Lsin = Math.Abs(swNetData.hs2 / Math.Sin(SMath.toRadians(180-swNetData.Hd2)));
                        double Lx = Lsin * Math.Cos(SMath.toRadians(180 - swNetData.Hd2));
                        double Ly = Lsin * Math.Sin(SMath.toRadians(180 - swNetData.Hd2));

                        swModel.SketchManager.CreateLine(swNetData.h, 0, 0, swNetData.h + Lx, Ly, 0);
                        add_LinName();
                        //swModel.Extension.SelectByID2(str_LinName(LineNames.Count - 1), "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                        //swModel.Extension.SelectByID2(str_LinName(0), "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
                        //swModel.AddDimension2(swNetData.h + Lx - (2E-02),0, 1E-02);
                        //tp_dim = (Dimension)swModel.Parameter("D2@草图1");
                        //tp_dim.Name = "HDelta2";
                        //tp_dim.SystemValue = SMath.toRadians(swNetData.hd2);

                        swModel.Extension.SelectByID2(str_LinName(LineNames.Count-1), "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                        if(swModel.AddVerticalDimension2(swNetData.h + (2E-02), 0, -1E-02)!=null)
                        {
                            tp_dim = (Dimension)swModel.Parameter("D2@草图1");
                            tp_dim.Name = "Hs2";
                            tp_dim.SystemValue = swNetData.hs2;
                        }

                
                        swModel.ClearSelection2(true);


                        swModel.Extension.SelectByID2("Point2", "SKETCHPOINT", 0, 0, 0, false, 0, null, 0);
                        swModel.SketchManager.CreateFillet(swNetData.d1, (int)swConstrainedCornerAction_e.swConstrainedCornerDeleteGeometry);
                    }
                    //草图绘制结束
                    swModel.SketchManager.InsertSketch(true);

                    //底基础网直径
                    swModel.Extension.SelectByID2("草图1", "SKETCH", 0, 0, 0, false, 4, null, 0);
                    swModel.FeatureManager.InsertProtrusionSwept4(false, false, 0, false, false, 0, 0, false, 0, 0, 0, 0, true, true, true, 0, true, true, swNetData.d1, 0);

                    swModel.Extension.SelectByID2("Point1@草图1", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);


                    ////选择参考点 创造基础面
                    //if (swNetData.hs1 != 0)
                    //{
                    //    swModel.Extension.SelectByID2("Point4@草图1", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
                    //    swModel.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, true, 1, null, 0);
                    //    swModel.FeatureManager.InsertRefPlane(4, 0, 1, 0, 0, 0);
                    //}


                    //创建草图平面
                    swModel.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, false, 0, null, 0);
                    swModel.SketchManager.InsertSketch(true);
                    double Xz = 0; 
                    foreach (DataGridViewRow item in dataGrid_View1.Rows)
                    {

                        if ((item.Cells[0].Value != null) && (item.Cells[1].Value != null))
                        {
                            for (int i = 0; i <int.Parse(item.Cells[1].Value.ToString()); i++)
                            {
                                Xz -= double.Parse(item.Cells[0].Value.ToString()) / 1000.00;

                                swModel.SketchManager.CreatePoint(-0, Xz, 0);
                            }
                        }
                    }

                    swModel.SketchManager.InsertSketch(true);
                    //swModel.EditRebuild3();
                    //Part.SketchManager.CreatePoint(-0#, -0.009412, 0#)
                    //草图阵列
                    swModel.Extension.SelectByID2("扫描1", "SOLIDBODY", 0, 0, 0, false, 256, null, 0);
                    swModel.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, true, 32, null, 0);
                    swModel.Extension.SelectByID2("草图4", "SKETCH", 0, 0, 0, true, 64, null, 0);
                    swModel.FeatureManager.FeatureSketchDrivenPattern(false, false);


                    swModel.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, true, 0, null, 0);
                    swModel.FeatureManager.InsertRefPlane(8, swNetData.d1 / 2, 0, 0, 0, 0);

                    //草图绘制基础圆
                    swModel.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, false, 0, null, 0);
                    swModel.SketchManager.InsertSketch(true);
                    ModelView modelView = (ModelView)swModel.ActiveView;
                    //swModel.ViewZoomtofit2();
                    modelView.Scale2 = 15.00;

                    double dy = (swNetData.d1 + swNetData.d2) / 2 - 0.1/1000.00;



                    double Xnoths1;
                    //swModel.Extension.SelectByID2("Line1@草图1", "EXTSKETCHSEGMENT", 0, 0, 0, true, 0, null, 1);
                    if (swNetData.hs1 != 0)
                    {
                        double offx = 0;
                        double offy = swNetData.hs1 - swNetData.d2 / 2;
                        if (swNetData.hd1 != 0)
                        {
                            double Lsin = Math.Abs(swNetData.hs1 / Math.Sin(SMath.toRadians(swNetData.Hd1)));
                            double Lx = Lsin * Math.Cos(SMath.toRadians(swNetData.Hd1));
                            double Ly = Lsin * Math.Sin(SMath.toRadians(swNetData.Hd1));
                            double x = -(swNetData.d1) / 2;
                            double y = -dy;
                            double xp = x * Math.Cos(SMath.toRadians(-swNetData.Hd1)) - y * Math.Sin(SMath.toRadians(-swNetData.Hd1));
                            double yp = x * Math.Sin(SMath.toRadians(-swNetData.Hd1)) + y * Math.Cos(SMath.toRadians(-swNetData.Hd1));

                            offx = Lx + xp;
                            offy = Ly>0?swNetData.hs1 - yp : -swNetData.hs1 - yp ;
                        }

                        double hs1x = offx;

                        swModel.SketchManager.CreateCircle(hs1x, offy, 0, hs1x + swNetData.d2 / 2, offy, 0);


                        Xnoths1 = -swNetData.d2 / 2;
                        //swModel.SketchManager.CreateCircle(Xnoths1, -dy, 0, Xnoths1 + swNetData.d2 / 2, -dy, 0);

                        foreach (DataGridViewRow item in dataGrid_View2.Rows)
                        {
                            if ((item.Cells[0].Value != null) && (item.Cells[1].Value != null))
                            {
                                for (int i = 0; i < int.Parse(item.Cells[1].Value.ToString()); i++)
                                {
                                    Xnoths1 += double.Parse(item.Cells[0].Value.ToString()) / 1000.00;
                                    if (Xnoths1 > (swNetData.h - (swNetData.d2 * 2))) continue;
                                    swModel.SketchManager.CreateCircle(Xnoths1, -dy, 0, Xnoths1 + swNetData.d2 / 2, -dy, 0);
                                }
                            }
                        }

                    }
                    else
                    {
                        Xnoths1 = swNetData.d2 / 2;
                        swModel.SketchManager.CreateCircle(Xnoths1, -dy, 0, Xnoths1 + swNetData.d2 / 2, -dy, 0);
                        modelView.RollBy(0);
                        foreach (DataGridViewRow item in dataGrid_View2.Rows)
                        {
                            if ((item.Cells[0].Value != null) && (item.Cells[1].Value != null))
                            {
                                for (int i = 0; i < int.Parse(item.Cells[1].Value.ToString()); i++)
                                {
                                    Xnoths1 += double.Parse(item.Cells[0].Value.ToString()) / 1000.00;
                                    if (Xnoths1 > (swNetData.h - (swNetData.d2 * 2))) continue;
                                    swModel.SketchManager.CreateCircle(Xnoths1, -dy, 0, Xnoths1 + swNetData.d2 / 2, -dy, 0);
                                }
                            }
                        }

                    }

                    if (swNetData.hs2 != 0)
                    {
                        double offx = 0;
                        double hs1x = 0;
                        double offy = swNetData.hs2 - swNetData.d2 / 2;
                        if (swNetData.hd2 != 0)
                        {
                            double Lsin = Math.Abs(swNetData.hs2/Math.Sin(SMath.toRadians(180-swNetData.Hd2)));
                            double Lx = Lsin * Math.Cos(SMath.toRadians(180 - swNetData.Hd2));
                            double Ly = Lsin * Math.Sin(SMath.toRadians(180 - swNetData.Hd2));
                            double x = (swNetData.d2) / 2;
                            double y = dy;
                            
                            double xp = x * Math.Cos(SMath.toRadians(- swNetData.Hd2))  - y * Math.Sin(SMath.toRadians(-swNetData.Hd2));
                            double yp = x * Math.Sin(SMath.toRadians(- swNetData.Hd2)) + y * Math.Cos(SMath.toRadians(-swNetData.Hd2));

                            offx = Lx + xp;
                            offy = (Ly > 0 ? swNetData.hs2 : -swNetData.hs2) + yp;
                        }

                        hs1x = swNetData.h + offx;

                        swModel.SketchManager.CreateCircle(hs1x, offy, 0, hs1x + swNetData.d2 / 2, offy, 0);
                    }
                    else
                    {

                        swModel.SketchManager.CreateCircle((swNetData.h-swNetData.d2 / 2), -dy, 0, swNetData.h, -dy, 0);
                    }

                    swModel.ClearSelection2(true);
                    //swModel.SketchManager.SketchUseEdge3(false, false);

                    //草图绘制结束
                    swModel.SketchManager.InsertSketch(true);
                    swModel.Extension.SelectByID2("草图5", "SKETCH", 0, 0, 0, false, 0, null, 0);
                    swModel.FeatureManager.FeatureExtrusion2(true, false, true, 1, 0, 0.01, 0.01, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);

                    swModel.ViewZoomtofit2();

                    string configName = "默认";
                    string databaseName = "SOLIDWORKS Materials";
                    string newPropName = "普通碳钢";

                    ((PartDoc)swModel).SetMaterialPropertyName2(configName, databaseName, newPropName);

                    ModelDocExtension swModelDocExt= swModel.Extension;

                    //写入属性
                    swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, InputDimFlag);
                    //CustomPropertyManager swCustProp = ((Configuration)swModel.GetActiveConfiguration()).CustomPropertyManager;
                    CustomPropertyManager swCustProp = (CustomPropertyManager)swModelDocExt.get_CustomPropertyManager("");
                    int bRet = swCustProp.Add3("SWNetData", (int)swCustomInfoType_e.swCustomInfoText, Conver_json(), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd); ;
                    //swCustProp.Set2("SWNetData","4625");

                    int massStatus = 0;
                    double[] massProperties;
                    massProperties = (double[])swModelDocExt.GetMassProperties(1, ref massStatus);

                    swModel.Extension.SelectByID2("草图4", "SKETCH", 0, 0, 0, false, 0, null, 0);
                    swModel.BlankSketch();
                    swModel.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, false, 0, null, 0);
                    swModel.BlankRefGeom();

                    swModel.ShowNamedView2("*等轴测", 7);
                    swModel.ViewZoomtofit2();
                    //Debug.Print(" Mass = " + massProperties[5]);
            

                    MessageBox.Show( $"重量:  {massProperties[5].ToString("0.0000")}  Kg\r\n"+
                        $"密度:   {massProperties[5] / massProperties[3]}  千克/立方米 ", $"质量信息{bRet}");


                }
            }

    


        }

        private void button_DJ_Click(object sender, EventArgs e)
        {
            if (dataGrid_View1.Rows[0].Cells[0]!=null)
            {
                if (int.Parse(dataGrid_View1.Rows[0].Cells[0].Value.ToString())== 0) return;
                dataGrid_View1.Rows[0].Cells[1].Value = (int)(swNetData.L1 / int.Parse(dataGrid_View1.Rows[0].Cells[0].Value.ToString()));
            }
        }

        private void button_DL_Click(object sender, EventArgs e)
        {
            if (dataGrid_View1.Rows[0].Cells[1] != null)
            {
                if (int.Parse(dataGrid_View1.Rows[0].Cells[1].Value.ToString()) == 0) return;
                dataGrid_View1.Rows[0].Cells[0].Value = (int)(swNetData.L1 / int.Parse(dataGrid_View1.Rows[0].Cells[1].Value.ToString()));
            }
        }

        private void button_DJx_Click(object sender, EventArgs e)
        {
            if (dataGrid_View2.Rows[0].Cells[0] != null)
            {
                if (int.Parse(dataGrid_View2.Rows[0].Cells[0].Value.ToString()) == 0) return;
                dataGrid_View2.Rows[0].Cells[1].Value = (int)(swNetData.H / int.Parse(dataGrid_View2.Rows[0].Cells[0].Value.ToString()));
            }
        }

        private void button_DLx_Click(object sender, EventArgs e)
        {
            if (dataGrid_View2.Rows[0].Cells[1] != null)
            {
                if (int.Parse(dataGrid_View2.Rows[0].Cells[1].Value.ToString()) == 0) return;
                dataGrid_View2.Rows[0].Cells[0].Value = (int)(swNetData.H / int.Parse(dataGrid_View2.Rows[0].Cells[1].Value.ToString()));
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start($"mailto: yang.linhao@hotmail.com");
        }

        private void dataGrid_View1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            double Ltmp = 0;
            try
            {
                foreach (DataGridViewRow item in dataGrid_View1.Rows)
                {
                    if ((item.Cells[0].Value != null) && (item.Cells[1].Value != null))
                    {
                        Ltmp += double.Parse(item.Cells[0].Value.ToString()) * int.Parse(item.Cells[1].Value.ToString());
                    }
                }
                Ltmp += swNetData.D1;

                //swNetData.L1 = Ltmp;
                this.tb_L1_t.Text = Ltmp.ToString();

            }
            catch (Exception)
            {

            }


        }


        private void dataGrid_View2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            double Htmp = 0;
            try
            {
                foreach (DataGridViewRow item in dataGrid_View2.Rows)
                {
                    if ((item.Cells[0].Value != null) && (item.Cells[1].Value != null))
                    {
                       Htmp += double.Parse(item.Cells[0].Value.ToString()) * int.Parse(item.Cells[1].Value.ToString());
                    }
                }


                //swNetData.L1 = Ltmp;
                //swNetData.H = Htmp;
                this.tb_H_t.Text = Htmp.ToString();

            }
            catch (Exception)
            {

            }
        }

        private void CreateNet_Load(object sender, EventArgs e)
        {

        }
        bool noreading=true;
        private void WTextChanged(object sender, EventArgs e)
        {

            if (noreading)
            {       
                dataGrid_View2_CellEndEdit(null, null);
                dataGrid_View1_CellEndEdit(null, null);
            }

            double Lhs1 = 0;
            double Lhs2 = 0;
            double cLs  = 1;

            try
            {
                swNetData.D1 = double.Parse(tB_D1.Text.ToString());
                swNetData.D2 = double.Parse(tB_D2.Text.ToString());
            }
            catch (Exception)
            {

            }
   

            if (swNetData.hd1 != 0 && swNetData.hs1 != 0)
            {
                Lhs1 = swNetData.hs1 / Math.Sin(SMath.toRadians(180 - swNetData.Hd1));
                cLs += 1;
            }

            if (swNetData.hd2 != 0 && swNetData.hs2 != 0)
            {
                Lhs2 = swNetData.hs2 / Math.Sin(SMath.toRadians(180 - swNetData.Hd2));
            }

            double wHs = (Lhs1 + Lhs2 + swNetData.h) * (swNetData.d1/2)* (swNetData.d1 / 2) *Math.PI * 7800;
            double cHs = 1;
            foreach (DataGridViewRow item in dataGrid_View1.Rows)
            {
                if ((item.Cells[0].Value != null) && (item.Cells[1].Value != null))
                {
                    try
                    {
                        cHs += int.Parse(item.Cells[1].Value.ToString());
                    }
                    catch (Exception)
                    {

                        continue;
                    }
                  
                }
            }
            double wH = wHs * cHs;

            double wLs = (swNetData.l1) * (swNetData.d2 / 2) * (swNetData.d2 / 2) * Math.PI * 7800;
          
            foreach (DataGridViewRow item in dataGrid_View2.Rows)
            {
                if ((item.Cells[0].Value != null) && (item.Cells[1].Value != null))
                {
                    cLs += int.Parse(item.Cells[1].Value.ToString());
                }
            }

            double wL = wLs * cLs;

            wL += wH;

            textBox_mess.Text=wL.ToString("0.000");

            if (tb_H_t.Text != tB_H.Text)
            {
                tB_H.BackColor = Color.Red;
            }
            else
            {
                tB_H.BackColor = Color.FromArgb(192, 255, 192);
            }

            if (tb_L1_t.Text != tB_L1.Text)
            {
                tB_L1.BackColor = Color.Red;
            }
            else
            {
                tB_L1.BackColor = Color.FromArgb(192, 255, 192);
            }

        }



        JsonClassData classData;
  
        public string Conver_json()
        {

            JsonClassData data=new JsonClassData();
            string JsonString = string.Empty;
            data.D1 = JsonConvert.SerializeObject(dt1);
            data.D2 = JsonConvert.SerializeObject(dt2);
            data.Net = JsonConvert.SerializeObject(swNetData);

            JsonString = JsonConvert.SerializeObject(data);


            return JsonString;
 
        }


        string SWNetData_valout = string.Empty;


        private void OpenSWNetFile_Click(object sender, EventArgs e)
        {
            string partDefault=string.Empty;
            string val = string.Empty;
       

            SldWorks swApp = ConnectToSolidWorks();
            if (swApp != null)
            {
                noreading=false;
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "SOLIDWORKS Part(*.SLDPRT)|*.SLDPRT"
                };

                openFileDialog.RestoreDirectory = true;
                openFileDialog.FilterIndex = 1;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    partDefault=openFileDialog.FileName;
                    var newDoc = swApp.NewDocument(partDefault, 0, 0, 0);

                    ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc; //当前零件
                    ModelDocExtension swModelDocExt = (ModelDocExtension)swModel.Extension;
                    CustomPropertyManager swCustProp = (CustomPropertyManager)swModelDocExt.get_CustomPropertyManager("");

                    bool status=swCustProp.Get4("SWNetData", false, out val, out SWNetData_valout);
                    if (status)
                    {
                        Debug.Print("Evaluated value:          " + SWNetData_valout);
                        classData = JsonConvert.DeserializeObject<JsonClassData>(SWNetData_valout);
                        try
                        {
                            dt1 = JsonConvert.DeserializeObject<DataTable>(classData.D1);
                            dt2 = JsonConvert.DeserializeObject<DataTable>(classData.D2);
                            swNetData.SetData(JsonConvert.DeserializeObject<SwNetData>(classData.Net));
                            dataGrid_View1.DataSource = dt1;
                            dataGrid_View2.DataSource = dt2;
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("读取失败",e.ToString());
                        }
                    

                    }
                 
                    
                }

            }
            noreading = true;
        }
    }
}
