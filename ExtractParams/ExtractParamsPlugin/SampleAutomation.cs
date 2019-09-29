/////////////////////////////////////////////////////////////////////
// Copyright (c) Autodesk, Inc. All rights reserved
// Written by Forge Partner Development
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
/////////////////////////////////////////////////////////////////////

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

using Inventor;
using Autodesk.Forge.DesignAutomation.Inventor.Utils;
using Newtonsoft.Json;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace ExtractParamsPlugin
{
  [ComVisible(true)]
  public class SampleAutomation
  {
    private readonly InventorServer inventorApplication;
    public bool IsLocalDebug = false;

    public SampleAutomation(InventorServer inventorApp)
    {
      inventorApplication = inventorApp;
    }

    public void Run(Document doc)
    {
      LogTrace("Run()");

      if (IsLocalDebug)
      {
        dynamic dynDoc = doc;
        string parameters = getParamsAsJson(dynDoc.ComponentDefinition.Parameters.UserParameters);

        System.IO.File.WriteAllText("documentParams.json", parameters);
      }
      else
      {
        string currentDir = System.IO.Directory.GetCurrentDirectory();
        LogTrace("Current Dir = " + currentDir);

        Dictionary<string, string> inputParameters = JsonConvert.DeserializeObject<Dictionary<string, string>>(System.IO.File.ReadAllText("inputParams.json"));
        logInputParameters(inputParameters);

        using (new HeartBeat())
        {
          if (inputParameters.ContainsKey("projectPath"))
          {
            string projectPath = inputParameters["projectPath"];
            string fullProjectPath = System.IO.Path.GetFullPath(System.IO.Path.Combine(currentDir, projectPath));
            LogTrace("fullProjectPath = " + fullProjectPath);
            DesignProject dp = inventorApplication.DesignProjectManager.DesignProjects.AddExisting(fullProjectPath);
            dp.Activate();
          }
          else
          {
            LogTrace("No 'projectPath' property");
          }

          if (inputParameters.ContainsKey("documentPath"))
          {
            string documentPath = inputParameters["documentPath"];

            string fullDocumentPath = System.IO.Path.GetFullPath(System.IO.Path.Combine(currentDir, documentPath));
            LogTrace("fullDocumentPath = " + fullDocumentPath);
            dynamic invDoc = inventorApplication.Documents.Open(fullDocumentPath);
            LogTrace("Opened input document file");
            dynamic compDef = invDoc.ComponentDefinition;

            string parameters = getParamsAsJson(compDef.Parameters.UserParameters);

            System.IO.File.WriteAllText("documentParams.json", parameters);
            LogTrace("Created documentParams.json");
          }
          else
          {
            LogTrace("No 'documentPath' property");
          }
        }
      }
    }

    public string getParamsAsJson(dynamic userParameters)
    {
      /* The resulting json will be like this:
        { 
          "length" : {
            "unit": "in",
            "value": "10 in",
            "values": ["5 in", "10 in", "15 in"]
          },
          "width": {
            "unit": "in",
            "value": "20 in",
          }
        }
      */
      List<object> parameters = new List<object>();
      foreach (dynamic param in userParameters)
      {
        List<object> paramProperties = new List<object>();
        if (param.ExpressionList != null)
        {
          string[] expressions = param.ExpressionList.GetExpressionList();
          JArray values = new JArray(expressions);
          paramProperties.Add(new JProperty("values", values));
        }
        paramProperties.Add(new JProperty("value", param.Expression));
        paramProperties.Add(new JProperty("unit", param.Units));

        parameters.Add(new JProperty(param.Name, new JObject(paramProperties.ToArray())));
      }
      JObject allParameters = new JObject(parameters.ToArray());
      string paramsJson = allParameters.ToString();
      LogTrace(paramsJson);

      return paramsJson;
    }

    public void logInputParameters(Dictionary<string, string> parameters)
    {
      foreach (KeyValuePair<string, string> entry in parameters)
      {
        try
        {
          LogTrace("Key = {0}, Value = {1}", entry.Key, entry.Value);
        }
        catch (Exception e)
        {
          LogTrace("Error with key {0}: {1}", entry.Key, e.Message);
        }
      }
    }

    #region Logging utilities

    /// <summary>
    /// Log message with 'trace' log level.
    /// </summary>
    private static void LogTrace(string format, params object[] args)
    {
      Trace.TraceInformation(format, args);
    }

    /// <summary>
    /// Log message with 'trace' log level.
    /// </summary>
    private static void LogTrace(string message)
    {
      Trace.TraceInformation(message);
    }

    /// <summary>
    /// Log message with 'error' log level.
    /// </summary>
    private static void LogError(string format, params object[] args)
    {
      Trace.TraceError(format, args);
    }

    /// <summary>
    /// Log message with 'error' log level.
    /// </summary>
    private static void LogError(string message)
    {
      Trace.TraceError(message);
    }

    #endregion
  }
}