using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using Sirenix.OdinInspector;
using Sirenix.OdinInspector.Editor;
using Sirenix.Utilities;
using Sirenix.Utilities.Editor;

[CreateAssetMenu(menuName = "Tools/Matching Weeks")]
public class MatchingWeeks : ScriptableObject
{
    public int iterationCount = 3;
    [FolderPath(AbsolutePath = true)]
    public string outputPath;
    public Entity_sheet_Modified referenceData;
    public List<Entity_sheet_Modified> weeks;

    [Button]
    public void CompareAndMatchWeeks() {

        int k = 0, countMerge = 0;
        do
        {
            if(k == 0) {
                countMerge = 1;
            } 
            else {
                countMerge = 0;
            }

            for (int i = 0; i < weeks.Count; i++)
            {
                weeks[i].referenceData = referenceData;
                countMerge += weeks[i].MatchingNamedGroupsBetweenSheets(applyLabeling : true);
            }
            ++k;
        }while(k < iterationCount && countMerge > 0);

        Debug.Log($"iteration: {k} \t countMerge: {countMerge}");
    }

    [Button]
    public void ExportData() {
        string temp, temp1;

        List<Entity_sheet_Modified> list = new List<Entity_sheet_Modified>();
        list.Add(referenceData);
        list.AddRange(weeks);

        for (int i = 0; i < list.Count; i++)
        {
            temp = list[i].output_GroupStatisticToCSV;
            temp1 = list[i].output_GroupToCSV;
            list[i].output_GroupStatisticToCSV = outputPath;
            list[i].output_GroupToCSV = outputPath;
            list[i].ExportGroupStatisticToCSV();
            list[i].ExportGroupToCSV();
            list[i].output_GroupStatisticToCSV = temp;
            list[i].output_GroupToCSV = temp1;
        }
    }
}
