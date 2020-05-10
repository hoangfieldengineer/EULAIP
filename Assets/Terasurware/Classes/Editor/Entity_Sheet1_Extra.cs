using System.Collections;
using System.Collections.Generic;
using System.Runtime.Versioning;
using Sirenix.OdinInspector;
using UnityEngine;
using UnityEngine.Serialization;

[System.Serializable]
public class Param_Extra : Entity_sheet1.Param
{
    [PropertyOrder(-1)] [TableColumnWidth(10)]
    public int group;
    
    [FoldoutGroup("Properties")] [FormerlySerializedAs("countRowsInGroup")]
    public int rows;
    [FormerlySerializedAs("userCount")] [FoldoutGroup("Properties")] // [FormerlySerializedAs("countUniqueUserIdInGroup")]
    public int users;
    [FormerlySerializedAs("countUniqueMachineIdInGroup")] [FoldoutGroup("Properties")]
    public int machines;
    [FormerlySerializedAs("countLicenseF")] [FoldoutGroup("Properties")]
    public int F;
    [FormerlySerializedAs("countLicenseH")] [FoldoutGroup("Properties")]
    public int H;

    public Param_Extra()
    {
        
    }
    
    public Param_Extra(Entity_sheet1.Param param) : base(param)
    {
        
    }

    public Param_Extra(Param_Extra extra) : base(extra)
    {
        this.@group = extra.@group;
        this.rows = extra.rows;
        this.users = extra.users;
        this.machines = extra.machines;
        this.F = extra.F;
        this.H = extra.H;
    }

    public override string ToString()
    {
        return $"{group},{rows},{users},{machines},{F},{H},{base.ToString()}";
    }
    
    public override object Clone()
    {
        return new Param_Extra(this);
    }
}


[CreateAssetMenu(menuName = "ScriptableObjects/EULA Sheet Extra")]
public class Entity_sheet1_Extra : ScriptableObject
{
    public List<Param_Extra> list = new List<Param_Extra>();

}