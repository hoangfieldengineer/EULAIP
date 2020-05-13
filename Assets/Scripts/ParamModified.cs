using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Sirenix.OdinInspector;
using Sirenix.OdinInspector.Editor;
using Sirenix.Utilities;
using Sirenix.Utilities.Editor;
using UnityEditor;
using UnityEngine;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

[System.SerializableAttribute]
public class ParamModified : Param_Extra
{

    [PropertyOrder(-3)]
    public string Description;
    [PropertyOrder(-2)]
    public string combinedGroup;

    public ParamModified(){
    }

    public ParamModified(Param_Extra extra) : base(extra)
    {
		
    }
	
    public ParamModified(ParamModified param) : base(param)
    {
        this.Description = param.Description;
        this.combinedGroup = param.combinedGroup;
    }

    public void FillDataIntoRow(HSSFRow row)
    {
        int i = 0;
        HSSFCell cell = (HSSFCell) row.CreateCell(i);
        cell.SetCellValue(Description);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(combinedGroup);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(@group);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(rows);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(users);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(machines);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(F);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(H);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(IP);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(LicenseHash);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(MachineId);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(Version);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(Location);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(Userid);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(count);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(SerialNumber);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(LicenseType);
        cell = (HSSFCell) row.CreateCell(++i);
        cell.SetCellValue(ShortVer);
    }
		
    public override string ToString()
    {
        return $"{Description},{combinedGroup},{base.ToString()}";
    }
	
    public override object Clone()
    {
        return new ParamModified(this);
    }
}