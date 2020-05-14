using System.Collections.Generic;
using Sirenix.OdinInspector;
using UnityEngine;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using UnityEngine.Serialization;


[System.Serializable]
public class Group
{
    public string name;
    
    [FoldoutGroup("Properties")]
    public string description;

    public string comment;
    
    [TableList(ShowIndexLabels = true)]
    public List<ParamModified> rows;
    
    [FormerlySerializedAs("countUniqueUserIdInGroup")] [FoldoutGroup("Properties")]
    public int users;
    [FormerlySerializedAs("countUniqueMachineIdInGroup")] [FoldoutGroup("Properties")]
    public int machines;
    [FormerlySerializedAs("countLicenseF")] [FoldoutGroup("Properties")]
    public int F;
    [FormerlySerializedAs("countLicenseH")] [FoldoutGroup("Properties")]
    public int H;
    

    public Group()
    {
        rows = new List<ParamModified>();
    }

    public override string ToString()
    {
        return $"[{name}]\tgroup: {rows[0].group}\tuser count: {users}\tmachine count: {machines}" +
               $"\tcount F: {F}\tcount H:{H}";
    }

    public float Compare(Group other, string sheet1, string sheet2, StringBuilder log, float loggingThreshold = 0.3f)
    {
        List<string> machines1 = rows.Select(x => x.MachineId).Where(x => !"#ERROR!".Equals(x)).Distinct().ToList();
        List<string> machines2 = other.rows.Select(x => x.MachineId).Where(x => !"#ERROR!".Equals(x)).Distinct().ToList();

        List<string> intersect = machines1.Intersect(machines2).ToList();

        float division = machines1.Count < machines2.Count ? machines1.Count : machines2.Count;
        float machineIntersectionPercent = intersect.Count * 1f / division;
	    
	    
        List<double> users1 = rows.Select(x => x.Userid).Distinct().ToList();
        List<double> users2 = other.rows.Select(x => x.Userid).Distinct().ToList();

        List<double> intersectUser = users1.Intersect(users2).ToList();

        division = users1.Count < users2.Count ? users1.Count : users2.Count;
        float userIntersectionPercent = intersectUser.Count * 1f / division;

        if (machineIntersectionPercent >= loggingThreshold || userIntersectionPercent >= loggingThreshold)
        {
            log.Clear();
            log.Append($"[{sheet1}] [{this.name}] vs [{other.name}] [{sheet2}]\n\n[{sheet1}] {this}\n\n[{sheet2}] {other}\n\nMachineID intersection: {machineIntersectionPercent}" +
                      $"\tUserID intersection: {userIntersectionPercent}\n\n");
        }
	    
        return machineIntersectionPercent > userIntersectionPercent ? machineIntersectionPercent : userIntersectionPercent;
    }

    public void CalculateStatistic()
    {
        machines = rows.Select(x => x.MachineId).Distinct().Count();
        users = rows.Select(x => x.Userid).Distinct().Count();
        F = rows.GroupBy(x => x.MachineId)
            .Select(g => new {machineId = g.Key, licenseType = g.FirstOrDefault().LicenseType}).ToList()
            .FindAll(x => x.licenseType.Equals("F")).Count;
        H = rows.GroupBy(x => x.MachineId)
            .Select(g => new {machineId = g.Key, licenseType = g.FirstOrDefault().LicenseType}).ToList()
            .FindAll(x => x.licenseType.Equals("H")).Count;

        IEnumerable<IGrouping<int, ParamModified>> grouping = rows.GroupBy(x => x.group);
        List<List<ParamModified>> groupedList = grouping.Select(grp => grp.ToList()).ToList();
        
        foreach (List<ParamModified> list in groupedList)
        {
            int rows = list.Count;
            int machines = list.Select(x => x.MachineId).Distinct().Count();
            int users = list.Select(x => x.Userid).Distinct().Count();
            int countF = list.GroupBy(x => x.MachineId)
                .Select(g => new {machineId = g.Key, licenseType = g.FirstOrDefault().LicenseType}).ToList()
                .FindAll(x => x.licenseType.Equals("F")).Count;
            int countH = list.GroupBy(x => x.MachineId)
                .Select(g => new {machineId = g.Key, licenseType = g.FirstOrDefault().LicenseType}).ToList()
                .FindAll(x => x.licenseType.Equals("H")).Count;

            for (int i = 0; i < list.Count; i++)
            {
                list[i].F = countF;
                list[i].H = countH;
                list[i].users = users;
                list[i].machines = machines;
                list[i].rows = rows;
            }
        }

    }

    public void FillDataIntoRow(HSSFRow row)
    {
        int i = 0;
        HSSFCell cell = (HSSFCell) row.CreateCell((short) i);
        cell.SetCellValue(name);
        cell = (HSSFCell) row.CreateCell((short) ++i);
        cell.SetCellValue(description);
        cell = (HSSFCell) row.CreateCell((short) ++i);
        cell.SetCellValue(comment);
        cell = (HSSFCell) row.CreateCell((short) ++i);
        cell.SetCellValue(rows[0].combinedGroup);
        cell = (HSSFCell) row.CreateCell((short) ++i);
        cell.SetCellValue(machines);
        cell = (HSSFCell) row.CreateCell((short) ++i);
        cell.SetCellValue(users);
        cell = (HSSFCell) row.CreateCell((short) ++i);
        cell.SetCellValue(F);
        cell = (HSSFCell) row.CreateCell((short) ++i);
        cell.SetCellValue(H);
        cell = (HSSFCell) row.CreateCell((short) ++i);
        cell.SetCellValue(F + H);
    }
}