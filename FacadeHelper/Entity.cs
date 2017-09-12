using AqlaSerializer;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using RvtDB = Autodesk.Revit.DB;

namespace FacadeHelper
{
    [SerializableType]
    public class CurtainPanelInfo : ElementInfoBase
    {
        private double _inf_Width_Metric;
        private double _inf_Height_Metric;
        //host
        private int _inf_HostId;
        private CurtainSystemInfo _inf_HostCurtainSystem;
        private WallInfo _inf_HostWall;
        //subs
        private List<ScheduleElementInfo> _inf_ScheduleElements;
        private List<DeepElementInfo> _inf_DeepElements;
        private List<GeneralElementInfo> _inf_GeneralElements;

        public double INF_Width_Metric { get { return _inf_Width_Metric; } set { _inf_Width_Metric = value; OnPropertyChanged(nameof(INF_Width_Metric)); } }
        public double INF_Height_Metric { get { return _inf_Height_Metric; } set { _inf_Height_Metric = value; OnPropertyChanged(nameof(INF_Height_Metric)); } }
        //host
        public int INF_HostId { get { return _inf_HostId; } set { _inf_HostId = value; OnPropertyChanged(nameof(INF_HostId)); } }
        public int FK_CurtainPanelInfo_CurtainSystemInfo { get; set; }
        public int FK_WallInfo_CurtainPanelInfo { get; set; }
        public CurtainSystemInfo INF_HostCurtainSystem { get { return _inf_HostCurtainSystem; } set { _inf_HostCurtainSystem = value; OnPropertyChanged(nameof(INF_HostCurtainSystem)); } }
        public WallInfo INF_HostWall { get { return _inf_HostWall; } set { _inf_HostWall = value; OnPropertyChanged(nameof(INF_HostWall)); } }
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
            #region GroupablePanel初始化
            INF_ElementId = p.Id.IntegerValue;
            INF_Name = p.Name;
            //INF_HostId = p.Host.Id.IntegerValue;
            //INF_Type = 51;

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
            //INF_System = p.Host.get_Parameter("立面系统").AsString();
            //INF_Direction = p.Host.get_Parameter("立面朝向").AsString();
            //INF_ZoneID = p.Host.get_Parameter("分区").AsInteger();

            INF_Width_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, p.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_WIDTH).AsDouble());
            INF_Height_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, p.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_HEIGHT).AsDouble());
            XYZ _xyzOrigin = p.GetTransform().Origin;
            INF_OriginX_US = _xyzOrigin.X;
            INF_OriginY_US = _xyzOrigin.Y;
            INF_OriginZ_US = _xyzOrigin.Z;
            INF_OriginX_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.X);
            INF_OriginY_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Y);
            INF_OriginZ_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Z);

            if ((_param = p.get_Parameter("分区区号")).HasValue)
            {
                INF_ZoneCode = _param.AsString();
                ResolveZoneCode();
            }
            else INF_ErrorInfo += $"[参数未设置：分区区号]";
            //立面楼层，分区区号 未读取
            #endregion
        }

        public override string ToString() { return INF_Code; }

        public void ResolveZoneCode()
        {
            //Z-00-99-AA-99
            var _array_field = INF_ZoneCode.Split('-');
            INF_System = _array_field[3].Substring(1, 1);
            INF_Direction = _array_field[3].Substring(0, 1);
            INF_Level = int.Parse(_array_field[2]);
            INF_ZoneID = int.Parse(_array_field[4]);
        }
    }

    [SerializableType]
    public class CurtainSystemInfo : ElementInfoBase
    {
        public List<CurtainPanelInfo> _inf_SubCurtainPanels;
        public int _inf_CSMode;
        public List<CurtainPanelInfo> INF_SubCurtainPanels { get { return _inf_SubCurtainPanels; } set { _inf_SubCurtainPanels = value; OnPropertyChanged(nameof(INF_SubCurtainPanels)); } }
        public int INF_CSMode { get { return _inf_CSMode; } set { _inf_CSMode = value; OnPropertyChanged(nameof(INF_CSMode)); } }// = -1;// 面板嵌板：1，立柱嵌板：2

        private CurtainSystem _cs;

        public CurtainSystemInfo() { INF_SubCurtainPanels = new List<CurtainPanelInfo>(); }
        public CurtainSystemInfo(CurtainSystem cs) : this()
        {
            #region GroupableCurtainSystem 初始化
            _cs = cs;
            INF_ElementId = cs.Id.IntegerValue;
            INF_Name = cs.Name;
            INF_CSMode = INF_Name.Contains("立柱幕墙") ? 2 : 1;

            INF_ErrorInfo = $"{cs.Id}";
            Parameter _param;
            if ((_param = cs.get_Parameter("立面系统")).HasValue) INF_System = _param.AsString(); else INF_ErrorInfo += $"[参数未设置：立面系统]";
            if ((_param = cs.get_Parameter("立面朝向")).HasValue) INF_Direction = _param.AsString(); else INF_ErrorInfo += $"[参数未设置：立面朝向]";
            if ((_param = cs.get_Parameter("分区")).HasValue) INF_ZoneID = _param.AsInteger(); else INF_ErrorInfo += $"[参数未设置：分区]";

            var _bxyz = cs.get_BoundingBox(cs.Document.ActiveView);
            INF_OriginX_US = _bxyz.Min.X;
            INF_OriginY_US = _bxyz.Min.Y;
            INF_OriginZ_US = _bxyz.Min.Z;
            INF_OriginX_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _bxyz.Min.X);
            INF_OriginY_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _bxyz.Min.Y);
            INF_OriginZ_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _bxyz.Min.Z);

            if (!INF_ErrorInfo.Contains("["))
            {
                INF_CS_ReadAllSubCurtainPanels(); //如果 系统|方位|子分区 已设置，读取子嵌板信息
            }
            #endregion
        }

        public void INF_CS_ReadAllSubCurtainPanels()
        {
            foreach (CurtainGrid _cgrid in _cs.CurtainGrids)
            {
                ICollection<ElementId> _panelids = _cgrid.GetPanelIds();
                foreach (ElementId eid in _panelids)
                {
                    RvtDB.Panel _p = _cs.Document.GetElement(eid) as RvtDB.Panel;
                    if ((_p as FamilyInstance).Symbol.Family.Name != "系统嵌板") INF_SubCurtainPanels.Add(new CurtainPanelInfo(_p));
                }

            }
        }
    }

    [SerializableType]
    public class WallInfo : ElementInfoBase
    {
        private List<CurtainPanelInfo> _inf_SubCurtainPanels;
        private int _inf_WallMode;
        public List<CurtainPanelInfo> INF_SubCurtainPanels { get { return _inf_SubCurtainPanels; } set { _inf_SubCurtainPanels = value; OnPropertyChanged(nameof(INF_SubCurtainPanels)); } }
        public int INF_WallMode { get { return _inf_WallMode; } set { _inf_WallMode = value; OnPropertyChanged(nameof(INF_WallMode)); } }// = -1;// 面板嵌板：1，立柱嵌板：2

        private Wall _wall;

        public WallInfo() { INF_SubCurtainPanels = new List<CurtainPanelInfo>(); }
        public WallInfo(Wall wall) : this()
        {
            _wall = wall;
            INF_ElementId = wall.Id.IntegerValue;
            INF_Name = wall.Name;
            INF_WallMode = INF_Name.Contains("立柱幕墙") ? 2 : 1;

            INF_ErrorInfo = $"{wall.Id}";
            Parameter _param;
            if ((_param = wall.get_Parameter("立面系统")).HasValue) INF_System = _param.AsString(); else INF_ErrorInfo += $"[参数未设置：立面系统]";
            if ((_param = wall.get_Parameter("立面朝向")).HasValue) INF_Direction = _param.AsString(); else INF_ErrorInfo += $"[参数未设置：立面朝向]";
            if ((_param = wall.get_Parameter("分区")).HasValue) INF_ZoneID = _param.AsInteger(); else INF_ErrorInfo += $"[参数未设置：分区]";

            var _bxyz = wall.get_BoundingBox(wall.Document.ActiveView);
            INF_OriginX_US = _bxyz.Min.X;
            INF_OriginY_US = _bxyz.Min.Y;
            INF_OriginZ_US = _bxyz.Min.Z;
            INF_OriginX_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _bxyz.Min.X);
            INF_OriginY_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _bxyz.Min.Y);
            INF_OriginZ_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _bxyz.Min.Z);

            if (!INF_ErrorInfo.Contains("["))
            {
                INF_Wall_ReadAllSubCurtainPanels(); //如果 系统|方位|子分区 已设置，读取子嵌板信息
            }

        }

        public void INF_Wall_ReadAllSubCurtainPanels()
        {
            ICollection<ElementId> _panelids = _wall.CurtainGrid.GetPanelIds();
            foreach (ElementId eid in _panelids)
            {
                RvtDB.Panel _p = _wall.Document.GetElement(eid) as RvtDB.Panel;
                if ((_p as FamilyInstance).Symbol.Family.Name != "系统嵌板") INF_SubCurtainPanels.Add(new CurtainPanelInfo(_p));
            }
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
    public abstract class ElementInfoBase: INotifyPropertyChanged
    {
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
        private ZoneInfoBase _inf_ZoneInfo;
        private int _inf_ZoneID;
        private string _inf_ZoneCode;
        private string _inf_Type_ZoneCode;
        private int _inf_Index;
        private string _inf_Code;
        private string _inf_Error_info;

        private uint _inf_TaskID;
        private int _inf_TaskLevel;
        private string _inf_TaskDuration;
        private uint _inf_TaskID_PreProcess;

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
        public ZoneInfoBase INF_ZoneInfo { get { return _inf_ZoneInfo; } set { _inf_ZoneInfo = value; OnPropertyChanged(nameof(INF_ZoneInfo)); } }
        public int INF_ZoneID { get { return _inf_ZoneID; } set { _inf_ZoneID = value; OnPropertyChanged(nameof(INF_ZoneID)); } }
        public string INF_ZoneCode { get { return _inf_ZoneCode; } set { _inf_ZoneCode = value; OnPropertyChanged(nameof(INF_ZoneCode)); } }
        public string INF_Type_ZoneCode { get { return _inf_Type_ZoneCode; } set { _inf_Type_ZoneCode = value; OnPropertyChanged(nameof(INF_Type_ZoneCode)); } }
        public int INF_Index { get { return _inf_Index; } set { _inf_Index = value; OnPropertyChanged(nameof(INF_Index)); } }
        public string INF_Code { get { return _inf_Code; } set { _inf_Code = value; OnPropertyChanged(nameof(INF_Code)); } }
        public string INF_ErrorInfo { get { return _inf_Error_info; } set { _inf_Error_info = value; OnPropertyChanged(nameof(INF_ErrorInfo)); } }

        public uint INF_TaskID { get { return _inf_TaskID; } set { _inf_TaskID = value; OnPropertyChanged(nameof(INF_TaskID)); } }
        public int INF_TaskLevel { get { return _inf_TaskLevel; } set { _inf_TaskLevel = value; OnPropertyChanged(nameof(INF_TaskLevel)); } }
        public string INF_TaskDuration { get { return _inf_TaskDuration; } set { _inf_TaskDuration = value; OnPropertyChanged(nameof(INF_TaskDuration)); } }
        public uint INF_TaskID_PreProcess { get { return _inf_TaskID_PreProcess; } set { _inf_TaskID_PreProcess = value; OnPropertyChanged(nameof(INF_TaskID_PreProcess)); } }

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
        public DesignZoneInfo(string zcode, DateTime zstart, DateTime zfinish) : base(zcode, zstart, zfinish) { }
    }

    [SerializableType]
    public class ProcessZoneInfo : ZoneInfoBase
    {
        public ProcessZoneInfo(string zcode) : base(zcode) { }
        public ProcessZoneInfo(string zcode, DateTime zstart, DateTime zfinish) : base(zcode, zstart, zfinish) { }
    }

    [SerializableType]
    public class ZoneInfoBase : INotifyPropertyChanged
    {
        //Z-00-99-AA-99, Z-01-99-AA-99, Z-02-99-AA-99
        private string _zoneCode;
        private DateTime _zoneStart;
        private DateTime _zoneFinish;
        private int _zoneDurationHours;

        private int _zoneType;
        private int _zoneLevel;
        private string _zoneDirection;
        private string _zoneSystem;
        private int _zoneIndex;

        private string _filterName;

        public string ZoneCode { get { return _zoneCode; } set { _zoneCode = value; OnPropertyChanged(nameof(ZoneCode)); } }
        public DateTime ZoneStart { get { return _zoneStart; } set { _zoneStart = value; OnPropertyChanged(nameof(ZoneStart)); } }
        public DateTime ZoneFinish { get { return _zoneFinish; } set { _zoneFinish = value; OnPropertyChanged(nameof(ZoneFinish)); } }
        public int ZoneDurationHours { get { return _zoneDurationHours; } set { _zoneDurationHours = value; OnPropertyChanged(nameof(ZoneDurationHours)); } }

        public int ZoneType { get { return _zoneType; } set { _zoneType = value; OnPropertyChanged(nameof(ZoneType)); } } //0,1,2
        public int ZoneLevel { get { return _zoneLevel; } set { _zoneLevel = value; OnPropertyChanged(nameof(ZoneLevel)); } }
        public string ZoneDirection { get { return _zoneDirection; } set { _zoneDirection = value; OnPropertyChanged(nameof(ZoneDirection)); } }
        public string ZoneSystem { get { return _zoneSystem; } set { _zoneSystem = value; OnPropertyChanged(nameof(ZoneSystem)); } }
        public int ZoneIndex { get { return _zoneIndex; } set { _zoneIndex = value; OnPropertyChanged(nameof(ZoneIndex)); } }

        public string FilterName { get { return _filterName; } set { _filterName = value; OnPropertyChanged(nameof(FilterName)); } }

        public ZoneInfoBase(string zcode)
        {
            ZoneCode = FilterName = zcode;
            string[] _segment_code = zcode.Split('-');
            ZoneType = int.Parse(_segment_code[1]);
            ZoneLevel = int.Parse(_segment_code[2]);
            ZoneDirection = _segment_code[3].Substring(0,1);
            ZoneSystem = _segment_code[3].Substring(1, 1);
            ZoneIndex = int.Parse(_segment_code[4]);

        }
        public ZoneInfoBase(string zcode, DateTime zstart, DateTime zfinish) : this(zcode)
        {
            ZoneStart = zstart;
            ZoneFinish = zfinish;
            ZoneDurationHours = (zfinish - zstart).Days * Global.OptionHoursPerDay + Global.OptionHoursPerDay;
        }

        public override string ToString() { return ZoneCode; }

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
