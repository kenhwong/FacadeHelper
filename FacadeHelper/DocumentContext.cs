using AqlaSerializer;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Interop;
//http://www.cnblogs.com/CreateMyself/p/6238455.html EntityFramework Core 1.1 Add、Attach、Update、Remove方法如何高效使用详解

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
        //private ObservableCollection<ZoneInfoBase> _zoneList = new ObservableCollection<ZoneInfoBase>();
        public List<ZoneLayerInfo> ZoneScheduleSimpleList { get; set; } = new List<ZoneLayerInfo>();
        public List<ZoneScheduleLayerInfo> ZoneScheduleLayerList { get; set; } = new List<ZoneScheduleLayerInfo>(); //Obsoleted property.
        public List<ZoneLayerInfo> ZoneLayerList { get; set; } = new List<ZoneLayerInfo>();

        public List<CurtainPanelInfo> CurtainPanelList { get; set; } = new List<CurtainPanelInfo>();
        public Dictionary<int, double> LevelDictionary { get; set; } = new Dictionary<int, double>();
        public List<LevelInfo> LevelList { get; set; } = new List<LevelInfo>();
        public List<ScheduleElementInfo> ScheduleElementList { get; set; } = new List<ScheduleElementInfo>();
        public List<DeepElementInfo> DeepElementList { get; set; } = new List<DeepElementInfo>();
        public ObservableCollection<ZoneInfoBase> ZoneList { get; set; } = new ObservableCollection<ZoneInfoBase>();
        public List<MullionInfo> MullionList { get; set; } = new List<MullionInfo>();

        [NonSerializableMember] public ILookup<string, CurtainPanelInfo> Lookup_CurtainPanels { get; set; }
        [NonSerializableMember] public ILookup<string, ScheduleElementInfo> Lookup_ScheduleElements { get; set; }
        [NonSerializableMember] public List<ParameterHelper.RawProjectParameterInfo> ParameterInfoList { get; set; } = new List<ParameterHelper.RawProjectParameterInfo>();

        public ZoneInfoBase CurrentZoneInfo = new ZoneInfoBase();

        //public SQLContext CurrentDBContext;
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public DocumentContent()
        {
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 1, EClassName = "玻璃", IsScheduled = true, ETaskLevel = 1, ETaskSubLevel = 22 });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 2, EClassName = "鋁件", IsScheduled = true, ETaskLevel = 2, ETaskSubLevel = 33 });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 3, EClassName = "鋁板", IsScheduled = true, ETaskLevel = 1, ETaskSubLevel = 23 });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 4, EClassName = "石材", IsScheduled = true, ETaskLevel = 1, ETaskSubLevel = 24 });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 5, EClassName = "鋼橫樑", IsScheduled = true, ETaskLevel = 0, ETaskSubLevel = 12 });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 6, EClassName = "鋼立柱", IsScheduled = true, ETaskLevel = 0, ETaskSubLevel = 11 });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 7, EClassName = "鋁橫樑", IsScheduled = true, ETaskLevel = 0, ETaskSubLevel = 24 });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 8, EClassName = "鋁立柱", IsScheduled = true, ETaskLevel = 0, ETaskSubLevel = 23 });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 9, EClassName = "緊固件", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 10, EClassName = "鐵板", IsScheduled = true, ETaskLevel = 1, ETaskSubLevel = 21});
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 11, EClassName = "保溫棉", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 12, EClassName = "膠", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 13, EClassName = "焊縫", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 14, EClassName = "鋼件", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 15, EClassName = "預埋件", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 16, EClassName = "連接板", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 17, EClassName = "钢线条", IsScheduled = true, ETaskLevel = 2, ETaskSubLevel = 31 });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 18, EClassName = "铝线条", IsScheduled = true, ETaskLevel = 2, ETaskSubLevel = 32 });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 21, EClassName = "窗扇", IsScheduled = true, ETaskLevel = 1, ETaskSubLevel = 27 });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 22, EClassName = "門扇", IsScheduled = true, ETaskLevel = 1, ETaskSubLevel = 28 });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 23, EClassName = "大五金", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 51, EClassName = "石材嵌板", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 52, EClassName = "玻璃嵌板", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 53, EClassName = "鋁板嵌板", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 54, EClassName = "百頁嵌板", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 61, EClassName = "立柱嵌板", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 71, EClassName = "连接组", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 81, EClassName = "金屬未歸類", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 82, EClassName = "非金屬未歸類", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 91, EClassName = "吊籃", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 92, EClassName = "汽吊", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 93, EClassName = "捲揚", IsScheduled = false });
            Global.ElementClassList.Add(new ElementClass { EClassIndex = 94, EClassName = "配電", IsScheduled = false });
            Global.TaskLevelClass = new int[][] { new int[] { 5, 6, 7, 8 }, new int[] { 1, 3, 4, 10, 21, 22 }, new int[] { 17, 18 } };
        }
    }

    [SerializableType]
    public class ZoneLayerInfo
    {
        public string HandleId { get; set; }
        public string ZoneCode { get; set; }
        public int ZoneLayer { get; set; }
        public DateTime ZoneStart { get; set; }
        public DateTime ZoneFinish { get; set; }
    }

    [SerializableType]
    public class ZoneScheduleLayerInfo
    {
        public string HandleId { get; set; }
        public string ZoneUniversalCode { get; set; }
        public DateTime[] ZoneLayerStart { get; set; } = new DateTime[3];
        public DateTime[] ZoneLayerFinish { get; set; } = new DateTime[3];
    }

    [SerializableType]
    public class ElementClass
    {
        public int EClassIndex { get; set; }
        public string EClassName { get; set; } = "Unset";
        public bool IsScheduled { get; set; } = false;
        public int ETaskLevel { get; set; } = -1;
        public int ETaskSubLevel { get; set; } = -1;
    }

    public class Global
    {
        public static DocumentContent DocContent;
        public static string DataFile;
        public static WindowInteropHelper winhelper;

        public static int OptionHoursPerDay = 8;
        public static int OptionDaysPerWeek = 7;

        public static List<ElementClass> ElementClassList = new List<ElementClass>();
        public static int[][] TaskLevelClass;
        
    }


}
