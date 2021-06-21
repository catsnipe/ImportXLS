using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class DebugTableDraw : MonoBehaviour
{
    // Start is called before the first frame update
    void Start()
    {
#if TRUE
        Debug.LogWarning("Debug display is possible by setting'#if FALSE'.");
#else
        ScriptableObject obj = Resources.Load<ScriptableObject>("Data_Sample");
        X_Sample.SetTable(obj);

        foreach (var row in X_Sample.Rows)
        {
            Debug.Log(
                $"ID:{row.ID}, " +
                $"No:{row.No}, " +
                $"Country:{row.CountryName}, " +
                $"City:{row.CityName}, " +
                $"Speeds: [{row.Speeds[0]} {row.Speeds[1]} {row.Speeds[2]}]"
            );
        }
#endif
    }
}
