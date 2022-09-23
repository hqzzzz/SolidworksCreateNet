using PSWStandalone;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

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


           

            DataColumn dt1_col1 = new DataColumn("间距",typeof(int));
            DataColumn dt1_col2 = new DataColumn("数量", typeof(int));

      

            DataColumn dt2_col1 = new DataColumn("间距", typeof(int));
            DataColumn dt2_col2 = new DataColumn("数量", typeof(int));

            dt1.Columns.Add(dt1_col1);
            dt1.Columns.Add(dt1_col2);
            DataRow row1 = dt1.NewRow();
            dt1.Rows.Add(row1);


            dt2.Columns.Add(dt2_col1);
            dt2.Columns.Add(dt2_col2);
            DataRow row2 = dt2.NewRow();
            dt2.Rows.Add(row2);

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

            dataGrid_View2.Rows[0].Cells[0].Value = "90";
            dataGrid_View2.Rows[0].Cells[1].Value = "10";
            dataGrid_View2.AutoGenerateColumns = true;
            dataGrid_View2.RowHeadersVisible = false;
            dataGrid_View2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGrid_View2.AllowUserToResizeColumns = false;
            dataGrid_View2.AllowUserToResizeRows = false;


            swNetData.D1 = 4;
            tB_D1.DataBindings.Add("Text", swNetData, "D1", false, DataSourceUpdateMode.OnPropertyChanged);

            swNetData.D2 = 3;
            tB_D2.DataBindings.Add("Text", swNetData, "D2", false, DataSourceUpdateMode.OnPropertyChanged);

            swNetData.H = 300;
            tB_H.DataBindings.Add("Text", swNetData, "H", false, DataSourceUpdateMode.OnPropertyChanged);
            swNetData.Hs1 = 30;
            tB_Hs1.DataBindings.Add("Text", swNetData, "Hs1", false, DataSourceUpdateMode.OnPropertyChanged);
            swNetData.Hs2 = 20;
            tB_Hs2.DataBindings.Add("Text", swNetData, "Hs2", false, DataSourceUpdateMode.OnPropertyChanged);
            swNetData.Hd1 = 100;
            tB_Hd1.DataBindings.Add("Text", swNetData, "Hd1", false, DataSourceUpdateMode.OnPropertyChanged);
            swNetData.Hd2 = 120;
            tB_Hd2.DataBindings.Add("Text", swNetData, "Hd2", false, DataSourceUpdateMode.OnPropertyChanged);
            swNetData.L1= 1000;
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
                        swModel.SketchManager.CreateLine(0, 0, 0, -0.003, swNetData.hs1, 0);
                        add_LinName();
                        swModel.Extension.SelectByID2(str_LinName(LineNames.Count-1), "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                        swModel.Extension.SelectByID2(str_LinName(0), "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
                        swModel.AddDimension2(2E-03, 0, -1E-03);
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
                        swModel.SketchManager.CreateLine(swNetData.h, 0, 0, swNetData.h + 0.003, swNetData.hs2, 0);
                        add_LinName();
                        swModel.Extension.SelectByID2(str_LinName(LineNames.Count-1), "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                        swModel.Extension.SelectByID2(str_LinName(0), "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
                        swModel.AddDimension2(swNetData.h-(2E-03), 0, -1E-03);
                        tp_dim = (Dimension)swModel.Parameter("D2@草图1");
                        tp_dim.Name = "HDelta2";
                        tp_dim.SystemValue = SMath.toRadians(swNetData.hd2);
                        
                        swModel.Extension.SelectByID2(str_LinName(LineNames.Count-1), "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
                        if(swModel.AddVerticalDimension2(swNetData.h + (2E-03), 0, -1E-03)!=null)
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


                    double Xnoths1 = 0;
                    ModelView modelView = (ModelView)swModel.ActiveView;
                    
                    modelView.Scale2 = 10.00;
                    
                    //swModel.Extension.SelectByID2("Line1@草图1", "EXTSKETCHSEGMENT", 0, 0, 0, true, 0, null, 1);
                    if (swNetData.hs1 != 0)
                    {
                        double offx = 0;
                        double hs1x = 0;
                        double offy = swNetData.hs1 - swNetData.d2 / 2;
                        if (swNetData.hd1!=0)
                        {
                            double Lsin = swNetData.hs1 / Math.Sin(SMath.toRadians(180 - swNetData.Hd1));
                            double Lx = Lsin * Math.Cos(SMath.toRadians(180 - swNetData.Hd1));
                            double x = (swNetData.d1) / 2;
                            double y = (swNetData.d1 + swNetData.d1) / 2.5;
                            double xp = x * Math.Cos(SMath.toRadians(-swNetData.Hd1)) - y * Math.Sin(SMath.toRadians(-swNetData.Hd1));
                            double yp = x * Math.Sin(SMath.toRadians(-swNetData.Hd1)) + y * Math.Cos(SMath.toRadians(-swNetData.Hd1));

                            offx = Lx + xp;
                            offy = swNetData.hs1 + yp;
                        }

                        hs1x = 0 - offx;

                        swModel.SketchManager.CreateCircle(hs1x, offy, 0, hs1x + swNetData.d2 / 2, offy, 0);

                      
                        Xnoths1 = swNetData.d2 * 2;
                        swModel.SketchManager.CreateCircle(Xnoths1, -(swNetData.d1 + swNetData.d2) / 2.1, 0, Xnoths1 + swNetData.d2 / 2, -(swNetData.d1 + swNetData.d2) / 2.1, 0);

                        foreach (DataGridViewRow item in dataGrid_View2.Rows)
                        {
                            if ((item.Cells[0].Value != null) && (item.Cells[1].Value != null))
                            {
                                for (int i = 0; i < int.Parse(item.Cells[1].Value.ToString()); i++)
                                {
                                    Xnoths1 += double.Parse(item.Cells[0].Value.ToString()) / 1000.00;
                                    if (Xnoths1 > (swNetData.h - swNetData.d2)) break;
                                    swModel.SketchManager.CreateCircle(Xnoths1, -(swNetData.d1 + swNetData.d2) / 2.1, 0, Xnoths1 + swNetData.d2 / 2, -(swNetData.d1 + swNetData.d2) / 2.1, 0);
                                }
                            }
                        }

                    }
                    else
                    {
                        Xnoths1 = swNetData.d2 / 2;
                       
                       
                        swModel.SketchManager.CreateCircle(Xnoths1, -(swNetData.d1 + swNetData.d2)/2.1, 0, Xnoths1 + swNetData.d2 / 2, -(swNetData.d1 + swNetData.d2)/2.1, 0);
                        modelView.RollBy(0);
                        foreach (DataGridViewRow item in dataGrid_View2.Rows)
                        {
                            if ((item.Cells[0].Value != null) && (item.Cells[1].Value != null))
                            {
                                for (int i = 0; i < int.Parse(item.Cells[1].Value.ToString()); i++)
                                {
                                    Xnoths1 += double.Parse(item.Cells[0].Value.ToString()) / 1000.00;
                                    if (Xnoths1 > (swNetData.h- swNetData.d2)) break;
                                    swModel.SketchManager.CreateCircle(Xnoths1, -(swNetData.d1 + swNetData.d2) / 2.1, 0, Xnoths1+swNetData.d2/2, -(swNetData.d1 + swNetData.d2) / 2.1, 0);
                                }
                            }
                        }
                      
                    }

                    if (swNetData.hs2 != 0)
                    {
                        double offx = 0;
                        double hs1x = 0;
                        double offy = swNetData.hs2 - swNetData.d2 / 2;
                        if (swNetData.hd1 != 0)
                        {
                            double Lsin = swNetData.hs2/Math.Sin(SMath.toRadians(180 - swNetData.Hd2));
                            double Lx = Lsin * Math.Cos(SMath.toRadians(180 - swNetData.Hd2));
                            
                            double x = (swNetData.d2) / 2;
                            double y = (swNetData.d1 + swNetData.d2) / 2.5;
                            
                            double xp = x * Math.Cos(SMath.toRadians(- swNetData.Hd2))  - y * Math.Sin(SMath.toRadians(-swNetData.Hd2));
                            double yp = x * Math.Sin(SMath.toRadians(- swNetData.Hd2)) + y * Math.Cos(SMath.toRadians(-swNetData.Hd2));

                            offx = Lx+xp;
                            offy = swNetData.hs2 + yp;
                        }

                        hs1x = swNetData.h + offx;

                        swModel.SketchManager.CreateCircle(hs1x, offy, 0, hs1x + swNetData.d2 / 2, offy, 0);
                    }
                    else
                    {

                        swModel.SketchManager.CreateCircle((swNetData.h-swNetData.d2 / 2), (swNetData.d1 + swNetData.d2) / 2.1, 0, swNetData.h, (swNetData.d1 + swNetData.d2) / 2.1, 0);
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

                    int massStatus = 0;
                    double[] massProperties;
                    massProperties = (double[])swModelDocExt.GetMassProperties(1, ref massStatus);

                    Debug.Print(" Mass = " + massProperties[5]);

                    MessageBox.Show( $"重量 : {massProperties[5].ToString("0.0000")} Kg\r\n 密度 :{massProperties[5] / massProperties[3]} 千克/立方米 ", "质量信息");


                }
            }

            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, InputDimFlag);


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
    }
}
