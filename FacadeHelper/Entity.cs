using AqlaSerializer;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using RvtDB = Autodesk.Revit.DB;

namespace FacadeHelper
{
    [SerializableType]
    public class CurtainPanelInfo : ElementInfoBase
    {
        private double _inf_Width_Metric;
        private double _inf_Height_Metric;
        //subs
        private List<ScheduleElementInfo> _inf_ScheduleElements;
        private List<DeepElementInfo> _inf_DeepElements;
        private List<GeneralElementInfo> _inf_GeneralElements;

        public double INF_Width_Metric { get { return _inf_Width_Metric; } set { _inf_Width_Metric = value; OnPropertyChanged(nameof(INF_Width_Metric)); } }
        public double INF_Height_Metric { get { return _inf_Height_Metric; } set { _inf_Height_Metric = value; OnPropertyChanged(nameof(INF_Height_Metric)); } }
        //subs
        public List<ScheduleElementInfo> INF_ScheduleElements { get { return _inf_ScheduleElements; } set { _inf_ScheduleElements = value; OnPropertyChanged(nameof(INF_ScheduleElements)); } }
        public List<DeepElementInfo> INF_DeepElements { get { return _inf_DeepElements; } set { _inf_DeepElements = value; OnPropertyChanged(nameof(INF_DeepElements)); } }
        public List<GeneralElementInfo> INF_GeneralElements { get { return _inf_GeneralElements; } set { _inf_GeneralElements = value; OnPropertyChanged(nameof(INF_GeneralElements)); } }

        public CurtainPanelInfo()
        {
            INF_Width_Metric = -1;
            INF_Height_Metric = -1;
            //INF_HostId = -1;
            INF_ScheduleElements = new List<ScheduleElementInfo>();
            INF_DeepElements = new List<DeepElementInfo>();
            INF_GeneralElements = new List<GeneralElementInfo>();
        }
        public CurtainPanelInfo(Panel p) : this()
        {
            #region CurtainPanelInfo 初始化
            INF_ElementId = p.Id.IntegerValue;
            INF_Name = p.Name;
            INF_ErrorInfo = $"{p.Id}";
            Parameter _param;
            if ((_param = p.get_Parameter("构件分项")).HasValue) INF_Type = _param.AsInteger();
            else
            {
                if (INF_Name.Contains("石材")) INF_Type = 51;
                if (INF_Name.Contains("玻璃")) INF_Type = 52;
                if (INF_Name.Contains("铝板")) INF_Type = 53;
                if (INF_Name.Contains("百页")) INF_Type = 54;
                if (INF_Name.Contains("立柱")) INF_Type = 61;
            }

            INF_Width_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, p.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_WIDTH).AsDouble());
            INF_Height_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, p.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_HEIGHT).AsDouble());
            XYZ _xyzOrigin = p.GetTransform().Origin;
            INF_OriginX_US = _xyzOrigin.X;
            INF_OriginY_US = _xyzOrigin.Y;
            INF_OriginZ_US = _xyzOrigin.Z;
            INF_OriginX_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.X);
            INF_OriginY_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Y);
            INF_OriginZ_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Z);

            /**
            if ((_param = p.get_Parameter("分区区号")).HasValue)
            {
                INF_ZoneCode = _param.AsString();
                ResolveZoneCode();
            }
            else INF_ErrorInfo += $"[参数未设置：分区区号]";
            **/
            #endregion
        }
        public CurtainPanelInfo(Panel p, string zonecode) : this(p)
        {
            if (Regex.IsMatch(zonecode, @"Z-00-\d{2}-[a-z|A-Z]{2}-\d{2}"))
            {
                INF_ZoneCode = zonecode.ToUpper();
                //Z-00-99-AA-99
                var _array_field = INF_ZoneCode.Split('-');
                //INF_ZoneLayer = int.Parse(_array_field[1]);
                INF_System = _array_field[3].Substring(1, 1);
                INF_Direction = _array_field[3].Substring(0, 1);
                INF_Level = int.Parse(_array_field[2]);
                INF_ZoneIndex = int.Parse(_array_field[4]);
            }
        }

        public override string ToString() => INF_Code;

    }

    [SerializableType]
    public class MullionInfo : ElementInfoBase
    {
        private double _inf_MullionLength_Metric = 0;
        public double INF_MullionLength_Metric { get { return _inf_MullionLength_Metric; } set { _inf_MullionLength_Metric = value; OnPropertyChanged(nameof(INF_MullionLength_Metric)); } }
        public MullionInfo() => INF_Type = 1;
        /// <summary>
        /// 按竪梃方向初始化
        /// </summary>
        /// <param name="mulliontype">竪梃方向，1: 立柱，2：橫樑</param>
        public MullionInfo(int mulliontype) => INF_Type = mulliontype;
        public MullionInfo(Mullion mu)
        {
            #region MullionInfo 初始化
            INF_ElementId = mu.Id.IntegerValue;
            INF_Name = mu.Name;
            INF_Type = 0;
            if (INF_Name.StartsWith("H")) INF_Type = 7;
            if (INF_Name.StartsWith("V")) INF_Type = 8;
            if (INF_Name.StartsWith("S")) INF_Type = 7;

            INF_ErrorInfo = $"{mu.Id}";

            INF_MullionLength_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, mu.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble());
            XYZ _xyzOrigin = mu.GetTransform().Origin;
            INF_OriginX_US = _xyzOrigin.X;
            INF_OriginY_US = _xyzOrigin.Y;
            INF_OriginZ_US = _xyzOrigin.Z;
            INF_OriginX_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.X);
            INF_OriginY_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Y);
            INF_OriginZ_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Z);


            Parameter _param;
            if ((_param = mu.get_Parameter("分区区号")).HasValue)
            {
                INF_ZoneCode = _param.AsString();
                ResolveZoneCode();
            }
            else INF_ErrorInfo += $"[参数未设置：分区区号]";
            //立面楼层，分区区号 未读取
            #endregion
        }
        public MullionInfo(Mullion mu, int mulliontype) : this(mu) => INF_Type = mulliontype;

        public override string ToString() => INF_Code;

        public void ResolveZoneCode()
        {
            INF_ZoneCode = INF_ZoneCode.ToUpper();
            //Z-00-99-AA-99
            var _array_field = INF_ZoneCode.Split('-');
            INF_ZoneLayer = int.Parse(_array_field[1]);
            INF_System = _array_field[3].Substring(1, 1);
            INF_Direction = _array_field[3].Substring(0, 1);
            INF_Level = int.Parse(_array_field[2]);
            INF_ZoneIndex = int.Parse(_array_field[4]);
        }

    }

    [SerializableType]
    public class ScheduleElementInfo : GeneralElementInfo
    {
        private bool _inf_HasDeepElements;
        private List<DeepElementInfo> _inf_DeepElements;
        public int FK_ScheduleElementInfo_CurtainPanelInfo { get; set; }
        public bool INF_HasDeepElements { get { return _inf_HasDeepElements; } set { _inf_HasDeepElements = value; OnPropertyChanged(nameof(INF_HasDeepElements)); } }
        public List<DeepElementInfo> INF_DeepElements { get { return _inf_DeepElements; } set { _inf_DeepElements = value; OnPropertyChanged(nameof(INF_DeepElements)); } }

        public ScheduleElementInfo() { INF_HasDeepElements = false; INF_DeepElements = new List<DeepElementInfo>(); }
    }

    [SerializableType]
    public class DeepElementInfo : GeneralElementInfo
    {
        private ScheduleElementInfo _inf_HostScheduleElement;
        public int FK_DeepElementInfo_ScheduleElementInfo { get; set; }
        public int FK_DeepElementInfo_CurtainPanelInfo { get; set; }
        public ScheduleElementInfo INF_HostScheduleElement { get { return _inf_HostScheduleElement; } set { _inf_HostScheduleElement = value; OnPropertyChanged(nameof(INF_HostScheduleElement)); } }
    }

    [SerializableType]
    public abstract class GeneralElementInfo : ElementInfoBase
    {
        private CurtainPanelInfo _inf_HostCurtainPanel;
        public int FK_GeneralElementInfo_CurtainPanelInfo { get; set; }
        public CurtainPanelInfo INF_HostCurtainPanel { get { return _inf_HostCurtainPanel; } set { _inf_HostCurtainPanel = value; OnPropertyChanged(nameof(INF_HostCurtainPanel)); } }
    }

    [SerializableType]
    public abstract class ElementInfoBase : INotifyPropertyChanged
    {
        private int _inf_ExternalLinkId = 0;
        private int _inf_ElementId;
        private string _inf_Name;
        private string _inf_System;
        private string _inf_Direction;
        private int _inf_Type;
        private int _inf_Level;
        private int _inf_GroupID;
        private double _inf_OriginX_Metric;
        private double _inf_OriginY_Metric;
        private double _inf_OriginZ_Metric;
        private double _inf_OriginX_US;
        private double _inf_OriginY_US;
        private double _inf_OriginZ_US;
        private ZoneInfoBase _inf_HostZoneInfo;
        private int _inf_ZoneIndex;
        private string _inf_ZoneCode;
        private int _inf_ZoneLayer;
        private int _inf_Index;
        private string _inf_Code;
        private string _inf_Error_info;
        private bool _inf_IsScheduled;

        private uint _inf_TaskID;
        private int _inf_TaskLayer = -1;
        private int _inf_TaskSubLayer = -1;
        private DateTime _inf_TaskStart;
        private DateTime _inf_TaskFinish;
        private string _inf_TaskDuration;
        private uint _inf_TaskID_PreProcess;

        public int INF_ExternalLinkId { get { return _inf_ExternalLinkId; } set { _inf_ExternalLinkId = value; OnPropertyChanged(nameof(INF_ExternalLinkId)); } }
        public int INF_ElementId { get { return _inf_ElementId; } set { _inf_ElementId = value; OnPropertyChanged(nameof(INF_ElementId)); } }
        public string INF_Name { get { return _inf_Name; } set { _inf_Name = value; OnPropertyChanged(nameof(INF_Name)); } }
        public string INF_System { get { return _inf_System; } set { _inf_System = value; OnPropertyChanged(nameof(INF_System)); } }
        public string INF_Direction { get { return _inf_Direction; } set { _inf_Direction = value; OnPropertyChanged(nameof(INF_Direction)); } }
        public int INF_Type { get { return _inf_Type; } set { _inf_Type = value; OnPropertyChanged(nameof(INF_Type)); } }
        public int INF_Level { get { return _inf_Level; } set { _inf_Level = value; OnPropertyChanged(nameof(INF_Level)); } }
        public int INF_GroupID { get { return _inf_GroupID; } set { _inf_GroupID = value; OnPropertyChanged(nameof(INF_GroupID)); } }
        public double INF_OriginX_Metric { get { return _inf_OriginX_Metric; } set { _inf_OriginX_Metric = value; OnPropertyChanged(nameof(INF_OriginX_Metric)); } }
        public double INF_OriginY_Metric { get { return _inf_OriginY_Metric; } set { _inf_OriginY_Metric = value; OnPropertyChanged(nameof(INF_OriginY_Metric)); } }
        public double INF_OriginZ_Metric { get { return _inf_OriginZ_Metric; } set { _inf_OriginZ_Metric = value; OnPropertyChanged(nameof(INF_OriginZ_Metric)); } }
        public double INF_OriginX_US { get { return _inf_OriginX_US; } set { _inf_OriginX_US = value; OnPropertyChanged(nameof(INF_OriginX_US)); } }
        public double INF_OriginY_US { get { return _inf_OriginY_US; } set { _inf_OriginY_US = value; OnPropertyChanged(nameof(INF_OriginY_US)); } }
        public double INF_OriginZ_US { get { return _inf_OriginZ_US; } set { _inf_OriginZ_US = value; OnPropertyChanged(nameof(INF_OriginZ_US)); } }
        public ZoneInfoBase INF_HostZoneInfo { get { return _inf_HostZoneInfo; } set { _inf_HostZoneInfo = value; OnPropertyChanged(nameof(INF_HostZoneInfo)); } }
        public int INF_ZoneIndex { get { return _inf_ZoneIndex; } set { _inf_ZoneIndex = value; OnPropertyChanged(nameof(INF_ZoneIndex)); } }
        public string INF_ZoneCode { get { return _inf_ZoneCode; } set { _inf_ZoneCode = value; OnPropertyChanged(nameof(INF_ZoneCode)); } }
        public int INF_ZoneLayer { get { return _inf_ZoneLayer; } set { _inf_ZoneLayer = value; OnPropertyChanged(nameof(INF_ZoneLayer)); } }
        public int INF_Index { get { return _inf_Index; } set { _inf_Index = value; OnPropertyChanged(nameof(INF_Index)); } }
        public string INF_Code { get { return _inf_Code; } set { _inf_Code = value; OnPropertyChanged(nameof(INF_Code)); } }
        public string INF_ErrorInfo { get { return _inf_Error_info; } set { _inf_Error_info = value; OnPropertyChanged(nameof(INF_ErrorInfo)); } }
        public bool INF_IsScheduled { get { return _inf_IsScheduled; } set { _inf_IsScheduled = value; OnPropertyChanged(nameof(INF_IsScheduled)); } }

        public uint INF_TaskID { get { return _inf_TaskID; } set { _inf_TaskID = value; OnPropertyChanged(nameof(INF_TaskID)); } }
        public int INF_TaskLayer { get { return _inf_TaskLayer; } set { _inf_TaskLayer = value; OnPropertyChanged(nameof(INF_TaskLayer)); } }
        public int INF_TaskSubLayer { get { return _inf_TaskSubLayer; } set { _inf_TaskSubLayer = value; OnPropertyChanged(nameof(INF_TaskSubLayer)); } }
        public DateTime INF_TaskStart { get { return _inf_TaskStart; } set { _inf_TaskStart = value; OnPropertyChanged(nameof(INF_TaskStart)); } }
        public DateTime INF_TaskFinish { get { return _inf_TaskFinish; } set { _inf_TaskFinish = value; OnPropertyChanged(nameof(INF_TaskFinish)); } }
        public string INF_TaskDuration { get { return _inf_TaskDuration; } set { _inf_TaskDuration = value; OnPropertyChanged(nameof(INF_TaskDuration)); } }
        public uint INF_TaskID_PreProcess { get { return _inf_TaskID_PreProcess; } set { _inf_TaskID_PreProcess = value; OnPropertyChanged(nameof(INF_TaskID_PreProcess)); } }

        public override string ToString() => INF_Code.ToString();

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    [SerializableType]
    public class LevelInfo : INotifyPropertyChanged
    {
        private int _inf_ElementId;
        private string _inf_LevelSign;
        private int _inf_LevelIndex;
        private double _inf_LevelElevation_Metric;
        private double _inf_LevelElevation_US;

        public int INF_ElementId { get { return _inf_ElementId; } set { _inf_ElementId = value; OnPropertyChanged(nameof(INF_ElementId)); } }
        public string INF_LevelSign { get { return _inf_LevelSign; } set { _inf_LevelSign = value; OnPropertyChanged(nameof(INF_LevelSign)); } }
        public int INF_LevelIndex { get { return _inf_LevelIndex; } set { _inf_LevelIndex = value; OnPropertyChanged(nameof(INF_LevelIndex)); } }
        public double INF_LevelElevation_Metric { get { return _inf_LevelElevation_Metric; } set { _inf_LevelElevation_Metric = value; OnPropertyChanged(nameof(INF_LevelElevation_Metric)); } }
        public double INF_LevelElevation_US { get { return _inf_LevelElevation_US; } set { _inf_LevelElevation_US = value; OnPropertyChanged(nameof(INF_LevelElevation_US)); } }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    [SerializableType]
    public class DesignZoneInfo : ZoneInfoBase
    {
        public DesignZoneInfo(string zcode) : base(zcode) { }
        //public DesignZoneInfo(string zcode, DateTime zstart, DateTime zfinish) : base(zcode, zstart, zfinish) { }
    }

    [SerializableType]
    public class ProcessZoneInfo : ZoneInfoBase
    {
        public ProcessZoneInfo(string zcode) : base(zcode) { }
        //public ProcessZoneInfo(string zcode, DateTime zstart, DateTime zfinish) : base(zcode, zstart, zfinish) { }
    }

    [SerializableType]
    public class ZoneInfoBase : INotifyPropertyChanged
    {
        //Z-00-99-AA-99, Z-01-99-AA-99, Z-02-99-AA-99
        private int _zoneExternalLinkId = 0;
        private string _zoneCode;

        private int _zoneLevel;
        private string _zoneDirection;
        private string _zoneSystem;
        private int _zoneIndex = 0;

        private string _filterName;

        public int ZoneExternalLinkId { get { return _zoneExternalLinkId; } set { _zoneExternalLinkId = value; OnPropertyChanged(nameof(ZoneExternalLinkId)); } }
        public string ZoneCode { get { return _zoneCode; } set { _zoneCode = value; OnPropertyChanged(nameof(ZoneCode)); } }

        public int ZoneLevel { get { return _zoneLevel; } set { _zoneLevel = value; OnPropertyChanged(nameof(ZoneLevel)); } }
        public string ZoneDirection { get { return _zoneDirection; } set { _zoneDirection = value; OnPropertyChanged(nameof(ZoneDirection)); } }
        public string ZoneSystem { get { return _zoneSystem; } set { _zoneSystem = value; OnPropertyChanged(nameof(ZoneSystem)); } }
        public int ZoneIndex { get { return _zoneIndex; } set { _zoneIndex = value; OnPropertyChanged(nameof(ZoneIndex)); } }

        public string FilterName { get { return _filterName; } set { _filterName = value; OnPropertyChanged(nameof(FilterName)); } }

        public ZoneInfoBase() { }
        public ZoneInfoBase(string zcode)
        {
            if (!Regex.IsMatch(zcode, @"Z-\d{2}-\d{2}-[S|N|E|W][A-Z]-\d{2}")) throw new FormatException("Invalid zone code format while Initializing ZoneInfoBase instance.");
            ZoneCode = FilterName = zcode;
            string[] _segment_code = zcode.Split('-');
            ZoneLevel = int.Parse(_segment_code[2]);
            ZoneDirection = _segment_code[3].Substring(0, 1);
            ZoneSystem = _segment_code[3].Substring(1, 1);
            ZoneIndex = int.Parse(_segment_code[4]);

        }

        public override string ToString() => ZoneCode;

        public SelectionFilterElement LoadFilter(UIDocument uidoc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(uidoc.Document);
            ICollection<Element> typecollection = collector.OfClass(typeof(SelectionFilterElement)).ToElements();

            SelectionFilterElement sef = typecollection.Cast<SelectionFilterElement>().FirstOrDefault(ele => ele.Name == FilterName);

            if (sef != null)
            {
                uidoc.Selection.Elements.Clear();
                foreach (ElementId id in sef.GetElementIds()) uidoc.Selection.Elements.Add(uidoc.Document.GetElement(id));
            }

            return sef;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
