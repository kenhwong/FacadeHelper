using AqlaSerializer;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Linq;
using System.Windows.Interop;
using System.Xml.Serialization;

namespace FacadeHelper
{
    /** SQLContext
    public class SQLContext : DbContext
    {
        public DbSet<CurtainSystemInfo> Element_CS { get; set; }
        public DbSet<WallInfo> Element_Wall { get; set; }
        public DbSet<CurtainPanelInfo> Element_Panel { get; set; }
        public DbSet<ScheduleElementInfo> Element_Schedule { get; set; }
        public DbSet<DeepElementInfo> Element_Deep { get; set; }
        public DbSet<GeneralElementInfo> Element_General { get; set; }
        public DbSet<LevelInfo> Levels { get; set; }

        public SQLContext(string connectionString): base(new SQLiteConnection() { ConnectionString = connectionString }, true) { }

        protected override void OnModelCreating(DbModelBuilder mb)
        {
            mb.Entity<CurtainSystemInfo>().ToTable("Element_CS").HasKey(m => m.INF_ElementId);
            mb.Entity<WallInfo>().ToTable("Element_Wall").HasKey(m => m.INF_ElementId);
            mb.Entity<CurtainPanelInfo>().ToTable("Element_Panel").HasKey(m => m.INF_ElementId);
            mb.Entity<ScheduleElementInfo>().ToTable("Element_Schedule");//.HasKey(m => m.INF_ElementId);
            mb.Entity<DeepElementInfo>().ToTable("Element_Deep");//.HasKey(m => m.INF_ElementId);
            mb.Entity<GeneralElementInfo>().ToTable("Element_General").HasKey(m => m.INF_ElementId);

            mb.Entity<ScheduleElementInfo>()
                .HasRequired(e => e.INF_HostCurtainPanel)
                .WithMany(p => p.INF_ScheduleElements)
                .HasForeignKey(e => e.FK_ScheduleElementInfo_CurtainPanelInfo);//FK_ScheduleElementInfo_CurtainPanelInfo 映射
            mb.Entity<DeepElementInfo>()
                .HasRequired(e => e.INF_HostCurtainPanel)
                .WithMany(p => p.INF_DeepElements)
                .HasForeignKey(e => e.FK_DeepElementInfo_CurtainPanelInfo);//FK_DeepElementInfoo_CurtainPanelInfo 映射
            mb.Entity<GeneralElementInfo>()
                .HasRequired(e => e.INF_HostCurtainPanel)
                .WithMany(p => p.INF_GeneralElements)
                .HasForeignKey(e => e.FK_GeneralElementInfo_CurtainPanelInfo);//FK_GeneralElementInfo_CurtainPanelInfo 映射
            mb.Entity<CurtainPanelInfo>()
                .HasRequired(p => p.INF_HostCurtainSystem)
                .WithMany(c => c.INF_SubCurtainPanels)
                .HasForeignKey(p => p.FK_CurtainPanelInfo_CurtainSystemInfo);//FK_CurtainPanelInfo_CurtainSystemInfo 映射
            mb.Entity<CurtainPanelInfo>()
                .HasRequired(p => p.INF_HostWall)
                .WithMany(c => c.INF_SubCurtainPanels)
                .HasForeignKey(p => p.FK_WallInfo_CurtainPanelInfo);//FK_WallInfo_CurtainPanelInfo 映射
            mb.Entity<DeepElementInfo>()
                .HasRequired(d => d.INF_HostScheduleElement)
                .WithMany(s => s.INF_DeepElements)
                .HasForeignKey(d => d.FK_DeepElementInfo_ScheduleElementInfo);//FK_DeepElementInfo_ScheduleElementInfo 映射

            mb.Entity<LevelInfo>()
                .ToTable("Level")
                .HasKey(m => m.INF_ElementId);
            mb.Entity<LevelInfo>()
                .Property(m => m.INF_LevelSign).IsRequired().HasMaxLength(8);

        }
    }
    **/

    public static class Constants
    {
        public const double Pi = 3.14159;
        public const double RVTPrecision = 0.001;
    }

    [SerializableType]
    public class DocumentContent// : INotifyPropertyChanged
    {
        public List<CurtainPanelInfo> CurtainPanelList { get; set; } = new List<CurtainPanelInfo>();
        public Dictionary<int, double> LevelDictionary { get; set; } = new Dictionary<int, double>();
        public List<LevelInfo> LevelList { get; set; } = new List<LevelInfo>();
        public List<ScheduleElementInfo> ScheduleElementList { get; set; } = new List<ScheduleElementInfo>();
        public List<DeepElementInfo> DeepElementList { get; set; } = new List<DeepElementInfo>();
        public ObservableCollection<ZoneInfoBase> ZoneList { get; set; } = new ObservableCollection<ZoneInfoBase>();
        public List<MullionInfo> MullionList { get; set; } = new List<MullionInfo>();

        public List<ExternalElementData> ExternalElementDataList { get; set; } = new List<ExternalElementData>();
        public List<CurtainPanelInfo> ExternalCurtainPanelList { get; set; } = new List<CurtainPanelInfo>();
        public List<ScheduleElementInfo> ExternalScheduleElementList { get; set; } = new List<ScheduleElementInfo>();
        public List<CurtainPanelInfo> FullCurtainPanelList { get; set; } = new List<CurtainPanelInfo>();
        public List<ScheduleElementInfo> FullScheduleElementList { get; set; } = new List<ScheduleElementInfo>();
        public ObservableCollection<ZoneInfoBase> FullZoneList { get; set; } = new ObservableCollection<ZoneInfoBase>();

        [NonSerializableMember] public List<ParameterHelper.RawProjectParameterInfo> ParameterInfoList { get; set; } = new List<ParameterHelper.RawProjectParameterInfo>();

        public ZoneInfoBase CurrentZoneInfo = new ZoneInfoBase();

        //public SQLContext CurrentDBContext;
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public DocumentContent()
        {
        }
    }

    [SerializableType]
    public class ExternalElementData
    {
        public int ExternalId { get; set; }
        public string ExternalFileName { get; set; }
        public List<CurtainPanelInfo> CurtainPanelList { get; set; } = new List<CurtainPanelInfo>();
        public List<ScheduleElementInfo> ScheduleElementList { get; set; } = new List<ScheduleElementInfo>();
        public List<ZoneInfoBase> ZoneList { get; set; } = new List<ZoneInfoBase>();
    }

    [Serializable, XmlRoot(Namespace = "", IsNullable = false)]
    [SerializableType]
    public class ZoneLayerInfo
    {
        [XmlAttribute] public string HandleId { get; set; }
        [XmlAttribute] public string ZoneCode { get; set; }
        [XmlAttribute] public int ZoneLayer { get; set; }
        [XmlAttribute] public DateTime ZoneStart { get; set; }
        [XmlAttribute] public DateTime ZoneFinish { get; set; }
        [XmlAttribute] public int ZoneDays { get; set; }
        [XmlAttribute] public int ZoneHours { get; set; }

        public ZoneLayerInfo() { }
    }

    [SerializableType]
    public class ZoneScheduleLayerInfo
    {
        public string HandleId { get; set; }
        public string ZoneUniversalCode { get; set; }
        public DateTime[] ZoneLayerStart { get; set; } = new DateTime[3];
        public DateTime[] ZoneLayerFinish { get; set; } = new DateTime[3];
    }

    [Serializable, XmlRoot(Namespace = "", IsNullable = false)]
    [SerializableType]
    public class ElementClass : INotifyPropertyChanged
    {
        private int _eClassIndex;
        private string _eClassName = "Unset";
        private bool _isScheduled = false;
        private int _eTaskLayer = -1;
        private int _eTaskSubLayer = -1;

        [XmlAttribute] public int EClassIndex { get { return _eClassIndex; } set { _eClassIndex = value; OnPropertyChanged(nameof(EClassIndex)); } }
        [XmlAttribute] public string EClassName { get { return _eClassName; } set { _eClassName = value; OnPropertyChanged(nameof(EClassName)); } }
        [XmlAttribute] public bool IsScheduled { get { return _isScheduled; } set { _isScheduled = value; OnPropertyChanged(nameof(IsScheduled)); } }
        [XmlAttribute] public int ETaskLayer { get { return _eTaskLayer; } set { _eTaskLayer = value; OnPropertyChanged(nameof(ETaskLayer)); } }
        [XmlAttribute] public int ETaskSubLayer { get { return _eTaskSubLayer; } set { _eTaskSubLayer = value; OnPropertyChanged(nameof(ETaskSubLayer)); } }

        public ElementClass()
        {
            EClassName = "Unset";
            IsScheduled = false;
            ETaskLayer = -1;
            ETaskSubLayer = -1;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    [Serializable, XmlRoot(Namespace = "", IsNullable = false)]
    [SerializableType]
    public class ElementIndexRange : INotifyPropertyChanged
    {
        private string _zoneCode = "";
        private int _elementType = 0;
        private int _indexMax = 0;

        [XmlAttribute] public string ZoneCode { get { return _zoneCode; } set { _zoneCode = value; OnPropertyChanged(nameof(ZoneCode)); } }
        [XmlAttribute] public int ElementType { get { return _elementType; } set { _elementType = value; OnPropertyChanged(nameof(ElementType)); } }
        [XmlAttribute] public int IndexMax { get { return _indexMax; } set { _indexMax = value; OnPropertyChanged(nameof(IndexMax)); } }

        public ElementIndexRange() { ZoneCode = "Unset"; ElementType = 0; IndexMax = 0; }
        public ElementIndexRange(string zonecode) : this() { ZoneCode = zonecode; }
        public ElementIndexRange(string zonecode, int etype) : this() { ZoneCode = zonecode; ElementType = etype; }
        public ElementIndexRange(string zonecode, int etype, int max) : this() { ZoneCode = zonecode; ElementType = etype; IndexMax = max; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    [Serializable, XmlRoot(Namespace = "", IsNullable = false)]
    [SerializableType]
    public class ElementFabricationInfo : INotifyPropertyChanged
    {
        int _elementType;
        [XmlAttribute] string FabrCode;
        [XmlAttribute] int FabrQuantity;
        [XmlAttribute] string IdentCode;

        [XmlAttribute] int ElementType;
        [XmlAttribute] string FabrCode;
        [XmlAttribute] int FabrQuantity;
        [XmlAttribute] string IdentCode;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public class Global
    {
        public static DocumentContent DocContent;
        public static string DataFile;
        public static WindowInteropHelper winhelper;

        public static int OptionHoursPerDay = 8;
        public static int OptionDaysPerWeek = 7;

        public static List<ElementIndexRange> ElementIndexRangeList = new List<ElementIndexRange>();
        public static List<ElementClass> ElementClassList = new List<ElementClass>();
        public static List<ZoneLayerInfo> ZoneLayerList = new List<ZoneLayerInfo>();
        public static int[][] TaskLevelClass;

        public static void UpdateAppConfig(string newKey, string newValue)
        {
            bool isModified = false;
            Configuration config = ConfigurationManager.OpenExeConfiguration(System.Reflection.Assembly.GetExecutingAssembly().Location);

            foreach (string key in config.AppSettings.Settings.AllKeys)
            {
                if (key == newKey)
                {
                    isModified = true;
                }
            }

            if (isModified) config.AppSettings.Settings.Remove(newKey);
            config.AppSettings.Settings.Add(newKey, newValue);
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        public static string GetAppConfig(string strKey)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(System.Reflection.Assembly.GetExecutingAssembly().Location);
            foreach (string key in config.AppSettings.Settings.AllKeys)
            {
                if (key == strKey)
                {
                    return config.AppSettings.Settings[strKey].Value;
                }
            }
            return null;
        }

        public static ElementIndexRange GetElementIndexRange(string zonecode, int etype)
        {
            ElementIndexRange eir = new ElementIndexRange(zonecode, etype);
            if (Global.ElementIndexRangeList?.Count == 0)
            {
                Global.ElementIndexRangeList.Add(eir);
                return eir;
            }
            else
            {
                return Global.ElementIndexRangeList.FirstOrDefault(r => r.ZoneCode.Equals(zonecode, StringComparison.CurrentCultureIgnoreCase) && r.ElementType == etype) ?? eir;
            }
        }

        public static void UpdateElementIndexRange(string zonecode, int etype, int max)
        {
            Global.ElementIndexRangeList.RemoveAll(r => r.ZoneCode.Equals(zonecode, StringComparison.CurrentCultureIgnoreCase) && r.ElementType == etype);
            Global.ElementIndexRangeList.Add(new ElementIndexRange(zonecode, etype, max));
        }
    }


}
