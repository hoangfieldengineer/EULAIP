using System;
using UnityEngine;
using System.Collections;
using System.Collections.Generic;
using Sirenix.OdinInspector;

[CreateAssetMenu(menuName = "ScriptableObjects/EULA Sheet")]
public class Entity_sheet1 : ScriptableObject
{	
	public List<Sheet> sheets = new List<Sheet> ();

	[System.SerializableAttribute]
	public class Sheet
	{
		public string name = string.Empty;
		public List<Param> list = new List<Param>();
	}

	[System.SerializableAttribute]
	public class Param : ICloneable
	{

		public string IP;
		public string LicenseHash;
		public string MachineId;
		public string Version;
		[TableColumnWidth(10)]
		public string Location;
		[TableColumnWidth(60)]
		public double Userid;
		[FoldoutGroup("Properties")]
		public int count;
		public string SerialNumber;
		public string LicenseType;
		[FoldoutGroup("Properties")] [TableColumnWidth(150)]
		public string ShortVer;

		public Param(){
		}

		public Param(Param param)
		{
			this.IP = param.IP;
			this.LicenseHash = param.LicenseHash;
			this.MachineId = param.MachineId;
			this.Version = param.Version;
			this.Location = param.Location;
			this.Userid = param.Userid;
			this.count = param.count;
			this.SerialNumber = param.SerialNumber;
			this.LicenseType = param.LicenseType;
			this.ShortVer = param.ShortVer;
		}
		
		public override string ToString()
		{
			return $"{IP},{LicenseHash},{MachineId},{Version},{Location},{Userid},{count},{SerialNumber},{LicenseType},{ShortVer}";
		}

		public virtual object Clone()
		{
			return new Param(this);
		}
	}
}

