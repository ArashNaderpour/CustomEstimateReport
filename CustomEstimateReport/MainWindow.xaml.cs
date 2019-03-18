using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Tekla.Structures.Model;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace CustomEstimateReport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Dictionary<String, Slab> slabData = new Dictionary<string, Slab>();
        Dictionary<String, Wall> wallData = new Dictionary<string, Wall>();
        Dictionary<String, Column> columnData = new Dictionary<string, Column>();
        Dictionary<String, Beam> beamData = new Dictionary<string, Beam>();
        Dictionary<String, Footing> footingData = new Dictionary<string, Footing>();
        Dictionary<String, ContinuousFooting> continuousFootingData = new Dictionary<string, ContinuousFooting>();
        Dictionary<String, Styrofoam> styrofoamData = new Dictionary<string, Styrofoam>();

        public MainWindow()
        {
            InitializeComponent();

            Model model = new Model();

            if (model.GetConnectionStatus() == false)
            {
                MessageBox.Show("Failed");
                this.Close();

                return;
            }

            else
            {
                ModelObjectEnumerator.AutoFetch = true;

                var allObjects = model.GetModelObjectSelector().GetAllObjects();

                IList<ModelObject> modelObjects = Methods.ToList(allObjects);

                foreach (ModelObject obj in modelObjects)
                {
                    if (obj.GetType().ToString() == "Tekla.Structures.Model.Assembly")
                    {
                        Hashtable nameValue = new Hashtable();

                        ArrayList nameObject = new ArrayList { "MAINPART.NAME" };

                        obj.GetStringReportProperties(nameObject, ref nameValue);

                        string assemblyName = nameValue[nameObject[0]].ToString();

                        string assemblyCategory = assemblyName.ToLower().Split('_')[0];

                        if (assemblyCategory == "slab")
                        {
                            Hashtable values = new Hashtable();
                            obj.GetAllReportProperties(SlabTemplate.stringProperties, SlabTemplate.doubleProperties, SlabTemplate.intProperties, ref values);

                            string name = values[SlabTemplate.stringProperties[0]].ToString();

                            string material;
                            try
                            {
                                material = values[SlabTemplate.stringProperties[1]].ToString();
                            }
                            catch
                            {
                                material = "Undefined";
                            }

                            double volumeGross;
                            try
                            {
                                volumeGross = Convert.ToDouble(values[SlabTemplate.doubleProperties[0]].ToString());
                            }
                            catch
                            {
                                volumeGross = 0;
                            }

                            double volumeNet;
                            try
                            {
                                volumeNet = Convert.ToDouble(values[SlabTemplate.doubleProperties[1]].ToString());
                            }
                            catch
                            {
                                volumeNet = 0;
                            }

                            double area;
                            try
                            {
                                area = Convert.ToDouble(values[SlabTemplate.doubleProperties[2]].ToString());
                            }
                            catch
                            {
                                area = 0;
                            }

                            double areaTop;
                            try
                            {
                                areaTop = Convert.ToDouble(values[SlabTemplate.doubleProperties[3]].ToString());
                            }
                            catch
                            {
                                areaTop = 0;
                            }

                            double areaBott;
                            try
                            {
                                areaBott = Convert.ToDouble(values[SlabTemplate.doubleProperties[4]].ToString());
                            }
                            catch
                            {
                                areaBott = 0;
                            }

                            double areaEdge = area - (areaTop + areaBott);

                            double perimeter = 0;
                            try
                            {
                                perimeter = Convert.ToDouble(values[SlabTemplate.doubleProperties[5]].ToString());
                            }
                            catch
                            {
                                MessageBox.Show(name);
                            }

                            if (slabData.ContainsKey(name))
                            {
                                slabData[name].volumeGross += Math.Round((volumeGross * 0.000000001307950619), 2);
                                slabData[name].volumeNet += Math.Round((volumeNet * 0.000000001307950619), 2);
                                slabData[name].areaTop += Math.Round((areaTop * 0.0000107639), 2);
                                slabData[name].areaBott += Math.Round((areaBott * 0.0000107639), 2);
                                slabData[name].areaEdge += Math.Round((areaEdge * 0.0000107639), 2); ;
                                slabData[name].perimeter += Math.Round((perimeter * 0.00328084), 2);
                                slabData[name].quantity += 1;
                            }

                            else
                            {
                                Slab slab = new Slab(name, material, volumeGross, volumeNet, areaEdge, areaTop, areaBott, perimeter, 1);

                                slabData.Add(name, slab);
                            }
                        }

                        else if (assemblyCategory == "wall" || assemblyCategory == "curb")
                        {
                            Hashtable values = new Hashtable();
                            obj.GetAllReportProperties(WallTemplate.stringProperties, WallTemplate.doubleProperties, WallTemplate.intProperties, ref values);

                            string name = values[WallTemplate.stringProperties[0]].ToString();

                            string material;
                            try
                            {
                                material = values[WallTemplate.stringProperties[1]].ToString();
                            }
                            catch
                            {
                                material = "Undefined";
                            }

                            double volumeGross;
                            try
                            {
                                volumeGross = Convert.ToDouble(values[WallTemplate.doubleProperties[0]].ToString());
                            }
                            catch
                            {
                                volumeGross = 0;
                            }

                            double volumeNet;
                            try
                            {
                                volumeNet = Convert.ToDouble(values[WallTemplate.doubleProperties[1]].ToString());
                            }
                            catch
                            {
                                volumeNet = 0;
                            }

                            double length;
                            try
                            {
                                length = Convert.ToDouble(values[WallTemplate.doubleProperties[2]].ToString());
                            }
                            catch
                            {
                                length = 0;
                            }

                            double areaTop;
                            try
                            {
                                areaTop = Convert.ToDouble(values[WallTemplate.doubleProperties[3]].ToString());
                            }
                            catch
                            {
                                areaTop = 0;
                            }

                            double areaEnd1;
                            try
                            {
                                areaEnd1 = Convert.ToDouble(values[WallTemplate.doubleProperties[4]].ToString());
                            }
                            catch
                            {
                                areaEnd1 = 0;
                            }

                            double areaEnd2;
                            try
                            {
                                areaEnd2 = Convert.ToDouble(values[WallTemplate.doubleProperties[5]].ToString());
                            }
                            catch
                            {
                                areaEnd2 = 0;
                            }

                            double areaSide1;
                            try
                            {
                                areaSide1 = Convert.ToDouble(values[WallTemplate.doubleProperties[6]].ToString());
                            }
                            catch
                            {
                                areaSide1 = 0;
                            }

                            double areaSide2;
                            try
                            {
                                areaSide2 = Convert.ToDouble(values[WallTemplate.doubleProperties[7]].ToString());
                            }
                            catch
                            {
                                areaSide2 = 0;
                            }

                            double areaSideGross;
                            try
                            {
                                areaSideGross = Convert.ToDouble(values[WallTemplate.doubleProperties[8]].ToString());
                            }
                            catch
                            {
                                areaSideGross = 0;
                            }

                            double areaSideNet;
                            try
                            {
                                areaSideNet = Convert.ToDouble(values[WallTemplate.doubleProperties[9]].ToString());
                            }
                            catch
                            {
                                areaSideNet = 0;
                            }

                            double areaOpening = areaSideGross - areaSideNet;

                            if (wallData.ContainsKey(name))
                            {
                                wallData[name].volumeGross += Math.Round((volumeGross * 0.000000001307950619), 2);
                                wallData[name].volumeNet += Math.Round((volumeNet * 0.000000001307950619), 2);
                                wallData[name].length += Math.Round((length * 0.00328084), 2);
                                wallData[name].areaTop += Math.Round((areaTop * 0.0000107639), 2);
                                wallData[name].areaEnd1 += Math.Round((areaEnd1 * 0.0000107639), 2);
                                wallData[name].areaEnd2 += Math.Round((areaEnd2 * 0.0000107639), 2);
                                wallData[name].areaSide1 += Math.Round((areaSide1 * 0.0000107639), 2);
                                wallData[name].areaSide2 += Math.Round((areaSide2 * 0.0000107639), 2);
                                wallData[name].areaSideGross += Math.Round((areaSideGross * 0.0000107639), 2);
                                wallData[name].areaOpening += Math.Round((areaOpening * 0.0000107639), 2);
                                wallData[name].quantity += 1;
                            }

                            else
                            {
                                Wall wall = new Wall(name, material, volumeGross, volumeNet, length, areaTop, areaEnd1, areaEnd2, areaSide1,
                                    areaSide2, areaSideGross, areaOpening, 1);

                                wallData.Add(name, wall);
                            }
                        }

                        else if (assemblyCategory == "column")
                        {
                            Hashtable values = new Hashtable();
                            obj.GetAllReportProperties(ColumnTemplate.stringProperties, ColumnTemplate.doubleProperties, ColumnTemplate.intProperties, ref values);

                            string name = values[ColumnTemplate.stringProperties[0]].ToString();

                            string material;
                            try
                            {
                                material = values[ColumnTemplate.stringProperties[1]].ToString();
                            }
                            catch
                            {
                                material = "Undefined";
                            }

                            double volumeGross;
                            try
                            {
                                volumeGross = Convert.ToDouble(values[ColumnTemplate.doubleProperties[0]].ToString());
                            }
                            catch
                            {
                                volumeGross = 0;
                            }

                            double height;
                            try
                            {
                                height = Convert.ToDouble(values[ColumnTemplate.doubleProperties[1]].ToString());
                            }
                            catch
                            {
                                height = 0;
                            }

                            double area;
                            try
                            {
                                area = Convert.ToDouble(values[ColumnTemplate.doubleProperties[2]].ToString());
                            }
                            catch
                            {
                                area = 0;
                            }

                            double areaTop;
                            try
                            {
                                areaTop = Convert.ToDouble(values[ColumnTemplate.doubleProperties[3]].ToString());
                            }
                            catch
                            {
                                areaTop = 0;
                            }

                            double areaBott;
                            try
                            {
                                areaBott = Convert.ToDouble(values[ColumnTemplate.doubleProperties[4]].ToString());
                            }
                            catch
                            {
                                areaBott = 0;
                            }

                            double areaSide = area - (areaBott + areaTop);

                            if (columnData.ContainsKey(name))
                            {
                                columnData[name].volumeGross += Math.Round((volumeGross * 0.000000001307950619), 2);
                                columnData[name].height += Math.Round((height * 0.00328084), 2);
                                columnData[name].areaSide += Math.Round((areaSide * 0.0000107639), 2);
                                columnData[name].quantity += 1;
                            }

                            else
                            {
                                Column column = new Column(name, material, volumeGross, height, areaSide, 1);

                                columnData.Add(name, column);
                            }
                        }

                        else if (assemblyCategory == "beam")
                        {
                            Hashtable values = new Hashtable();
                            obj.GetAllReportProperties(BeamTemplate.stringProperties, BeamTemplate.doubleProperties, BeamTemplate.intProperties, ref values);

                            string name = values[BeamTemplate.stringProperties[0]].ToString();

                            string material;
                            try
                            {
                                material = values[BeamTemplate.stringProperties[1]].ToString();
                            }
                            catch
                            {
                                material = "Undefined";
                            }

                            double volumeGross;
                            try
                            {
                                volumeGross = Convert.ToDouble(values[BeamTemplate.doubleProperties[0]].ToString());
                            }
                            catch
                            {
                                volumeGross = 0;
                            }

                            double length;
                            try
                            {
                                length = Convert.ToDouble(values[BeamTemplate.doubleProperties[1]].ToString());
                            }
                            catch
                            {
                                length = 0;
                            }

                            double width;
                            try
                            {
                                width = Convert.ToDouble(values[BeamTemplate.doubleProperties[2]].ToString());
                            }
                            catch
                            {
                                width = 0;
                            }

                            double areaBott;
                            try
                            {
                                areaBott = Convert.ToDouble(values[BeamTemplate.doubleProperties[3]].ToString());
                            }
                            catch
                            {
                                areaBott = 0;
                            }

                            // Converted to SQR-FOOT from the start
                            double areaSide = Math.Round((length * 0.00328084), 2) * Math.Round((width * 0.00328084), 2) * 2;

                            if (beamData.ContainsKey(name))
                            {
                                beamData[name].volumeGross += Math.Round((volumeGross * 0.000000001307950619), 2);
                                beamData[name].length += Math.Round((length * 0.00328084), 2);
                                beamData[name].areaBott += Math.Round((areaBott * 0.0000107639), 2);
                                beamData[name].areaSide += areaSide;
                                beamData[name].quantity += 1;
                            }

                            else
                            {
                                Beam beam = new Beam(name, material, volumeGross, length, areaBott, areaSide, 1);

                                beamData.Add(name, beam);
                            }
                        }

                        else if (assemblyCategory == "footing")
                        {
                            Hashtable values = new Hashtable();
                            obj.GetAllReportProperties(FootingTemplate.stringProperties, FootingTemplate.doubleProperties, FootingTemplate.intProperties, ref values);

                            string name = values[FootingTemplate.stringProperties[0]].ToString();

                            string material;
                            try
                            {
                                material = values[FootingTemplate.stringProperties[1]].ToString();
                            }
                            catch
                            {
                                material = "Undefined";
                            }

                            double volumeGross;
                            try
                            {
                                volumeGross = Convert.ToDouble(values[FootingTemplate.doubleProperties[0]].ToString());
                            }
                            catch
                            {
                                volumeGross = 0;
                            }

                            double area;
                            try
                            {
                                area = Convert.ToDouble(values[FootingTemplate.doubleProperties[1]].ToString());
                            }
                            catch
                            {
                                area = 0;
                            }

                            double areaTop;
                            try
                            {
                                areaTop = Convert.ToDouble(values[FootingTemplate.doubleProperties[2]].ToString());
                            }
                            catch
                            {
                                areaTop = 0;
                            }

                            double areaBott;
                            try
                            {
                                areaBott = Convert.ToDouble(values[FootingTemplate.doubleProperties[3]].ToString());
                            }
                            catch
                            {
                                areaBott = 0;
                            }

                            double areaEdge = area - (areaTop + areaBott);

                            if (footingData.ContainsKey(name))
                            {
                                footingData[name].volumeGross += Math.Round((volumeGross * 0.000000001307950619), 2);
                                footingData[name].areaTop += Math.Round((areaTop * 0.0000107639), 2);
                                footingData[name].areaBott += Math.Round((areaBott * 0.0000107639), 2);
                                footingData[name].areaEdge += Math.Round((areaEdge * 0.0000107639), 2); ;
                                footingData[name].quantity += 1;
                            }

                            else
                            {
                                Footing footing = new Footing(name, material, volumeGross, areaEdge, areaTop, areaBott, 1);

                                footingData.Add(name, footing);
                            }
                        }

                        else if (assemblyCategory == "continuous")
                        {
                            Hashtable values = new Hashtable();
                            obj.GetAllReportProperties(ContinuousFootingTemplate.stringProperties, ContinuousFootingTemplate.doubleProperties,
                                ContinuousFootingTemplate.intProperties, ref values);

                            string name = values[ContinuousFootingTemplate.stringProperties[0]].ToString();

                            string material;
                            try
                            {
                                material = values[ContinuousFootingTemplate.stringProperties[1]].ToString();
                            }
                            catch
                            {
                                material = "Undefined";
                            }

                            double volumeGross;
                            try
                            {
                                volumeGross = Convert.ToDouble(values[ContinuousFootingTemplate.doubleProperties[0]].ToString());
                            }
                            catch
                            {
                                volumeGross = 0;
                            }

                            double length;
                            try
                            {
                                length = Convert.ToDouble(values[ContinuousFootingTemplate.doubleProperties[1]].ToString());
                            }
                            catch
                            {
                                length = 0;
                            }

                            double areaTop;
                            try
                            {
                                areaTop = Convert.ToDouble(values[ContinuousFootingTemplate.doubleProperties[2]].ToString());
                            }
                            catch
                            {
                                areaTop = 0;
                            }

                            double areaEnd1;
                            try
                            {
                                areaEnd1 = Convert.ToDouble(values[ContinuousFootingTemplate.doubleProperties[3]].ToString());
                            }
                            catch
                            {
                                areaEnd1 = 0;
                            }

                            double areaEnd2;
                            try
                            {
                                areaEnd2 = Convert.ToDouble(values[ContinuousFootingTemplate.doubleProperties[4]].ToString());
                            }
                            catch
                            {
                                areaEnd2 = 0;
                            }

                            double areaSide1;
                            try
                            {
                                areaSide1 = Convert.ToDouble(values[ContinuousFootingTemplate.doubleProperties[5]].ToString());
                            }
                            catch
                            {
                                areaSide1 = 0;
                            }

                            double areaSide2;
                            try
                            {
                                areaSide2 = Convert.ToDouble(values[ContinuousFootingTemplate.doubleProperties[6]].ToString());
                            }
                            catch
                            {
                                areaSide2 = 0;
                            }

                            if (continuousFootingData.ContainsKey(name))
                            {
                                continuousFootingData[name].volumeGross += Math.Round((volumeGross * 0.000000001307950619), 2);
                                continuousFootingData[name].length += Math.Round((length * 0.00328084), 2);
                                continuousFootingData[name].areaTop += Math.Round((areaTop * 0.0000107639), 2);
                                continuousFootingData[name].areaEnd1 += Math.Round((areaEnd1 * 0.0000107639), 2);
                                continuousFootingData[name].areaEnd2 += Math.Round((areaEnd2 * 0.0000107639), 2);
                                continuousFootingData[name].areaSide1 += Math.Round((areaSide1 * 0.0000107639), 2);
                                continuousFootingData[name].areaSide2 += Math.Round((areaSide2 * 0.0000107639), 2);
                                continuousFootingData[name].quantity += 1;
                            }

                            else
                            {
                                ContinuousFooting continuousFooting = new ContinuousFooting(name, material, volumeGross, length, areaTop, areaEnd1, areaEnd2, areaSide1, areaSide2, 1);

                                continuousFootingData.Add(name, continuousFooting);
                            }
                        }

                        else if (assemblyCategory == "styrofoam")
                        {
                            Hashtable values = new Hashtable();
                            obj.GetAllReportProperties(StyrofoamTemplate.stringProperties, StyrofoamTemplate.doubleProperties, StyrofoamTemplate.intProperties, ref values);

                            string name = values[StyrofoamTemplate.stringProperties[0]].ToString();

                            string material = "Styrofoam";

                            double volumeGross;
                            try
                            {
                                volumeGross = Convert.ToDouble(values[StyrofoamTemplate.doubleProperties[0]].ToString());
                            }
                            catch
                            {
                                volumeGross = 0;
                            }

                            if (styrofoamData.ContainsKey(name))
                            {
                                styrofoamData[name].volumeGross += Math.Round((volumeGross * 0.000000001307950619), 2);
                                styrofoamData[name].quantity += 1;
                            }

                            else
                            {
                                Styrofoam styrofoam = new Styrofoam(name, material, volumeGross, 1);
                                styrofoamData.Add(name, styrofoam);
                            }
                        }
                    }
                }

                // Export Model Data To Excel
                try
                {
                    Methods.exportToExcel(slabData, beamData, columnData, footingData, wallData, continuousFootingData, styrofoamData);
                }
                catch (Exception err)
                {
                    this.label.Content = "FAILED!";
                }
            }
        }
    }
}
