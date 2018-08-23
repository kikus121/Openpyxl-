"""
Created By: Kamil Buczkowski
Last Modified: 24/07/2018

This file contains all the data structures used to in functions

"""

systemPrametersList = ['Name','CreatedBy','CreatedOn','Category','ComponentNames','Extsystem','ExtObject','ExtIdentifier','Description ']
componentPrametersList = ["Name","CreatedBy","CreatedOn","TypeName","Space","Description","Extsystem","ExtObject","ExtIdentifier","SerialNumber","InstallationDate","WarrantyStartDate","TagNumber","BarCode" ,"AssetIdentifier"]	 	 	 						 	 			
typePrametersList = ["Name","CreatedBy","CreatedOn","Category","Description","AssetType","Manufacturer","ModelNumber","WarrantyGuarantorParts","WarrantyDurationParts","WarrantyGuarantorLabor","WarrantyDurationLabor","WarrantyDurationUnit","ExtSystem","ExtObject","ExtIdentifier","ReplacementCost","ExpectedLife","WarrantyDescription","NominalLength","NominalWidth","NominalHeight"]
zonePrametersList = ["Name","CreatedBy","CreatedOn","Category","SpacesName","Extsystem","ExtObject","ExtIdentifier"]
spacePrametersList = ["Name","CreatedBy","CreatedOn","Category","FloorName","Description","Extsystem","ExtObject","ExtIdentifier","RoomTag","UsableHeight","GrossArea","NetArea"]
floorPrametersList = ["Name","CreatedBy","CreatedOn","Category","Extsystem","ExtObject","ExtIdentifier","Description","Elevation","Height"]
facilityPrametersList = ["Name","CreatedBy","CreatedOn","Category","ProjectName","SiteName","LinearUnits","AreaUnits","VolumeUnits","CurrencyUnit","AreaMeasurement","ExternalSystem","ExternalProjectObject","ExternalProjectIdentifier","ExternalSiteObject","ExternalSiteIdentifier","ExternalFacilityObject","ExternalFacilityIdentifier","Description","ProjectDescription","SiteDescription","Phase" ]
contactPrametersList = ["Email","CreatedBy","CreatedOn","Category","Company","Phone","ExtSystem","ExtObject","ExtIdentifier","Department","OrganizationCode","GivenName","FamilyName","Street","PostalBox","Town","StateRegion","PostalCode","Country"]
instructionPrametersList = ["Title","Version","Release","Status","Region"]

"""
Check Multiple Properties = 0 
Check Multiple Properties - m2 = 1
Check Multiple Properties - mm = 2
Check Multiple Properties for Related Element = 3
Check Omniclass Depth = 4
Check Property Exists = 5
Check Property Value Equals = 6
Check Property Value Equals Any of Multiple Values = 7
Check Property Value Exists = 8
Check Property Value Exists Against Quantity = 9
Check Property Value Is In Range = 10
Check Property Value is Unique = 11
Check Property Value Matches Pattern = 12
Ensure Property Missing = 13
Ensure Property Value Missing = 14
Ensure Property Value Not Equal = 15
Check Multiple Properties - 2 args = 16
"""
ruleList = [ "Check Multiple Properties","Check Multiple Properties - m2","Check Multiple Properties - mm",
             "Check Multiple Properties for Related Element","Check Omniclass Depth","Check Property Exists",
             "Check Property Value Equals","Check Property Value Equals Any of Multiple Values","Check Property Value Exists",
             "Check Property Value Exists Against Quantity","Check Property Value Is In Range","Check Property Value is Unique",
             "Check Property Value Matches Pattern","Ensure Property Missing","Ensure Property Value Missing",
             "Ensure Property Value Not Equal","Check Multiple Properties - 2 args"]

            
stageDic ={
    "1(i)":1,
    "2(ii)a":2,
    "2(ii)b":3,
    "3(iii)":4,
    "4(iv)":5,
    "5(v) ":6
    
    }

systemDic ={
    "Varibles": {
        "NumberOfRules": 0,
        "NumberOfParameters": 9
    },
    "ID": {
        "Classification": "",
        "Model": "",
        "Role":"",
        "Description": ""
    },
    "excelSheetLocation": {
        "CellStart":"CM106",
        "CellEnd":"CU108"
    },
    "Rules": {
        "Check Multiple Properties": bool(False) ,
        "Check Multiple Properties - m2": bool(False),
        "Check Multiple Properties - mm": bool(False),
        "Check Multiple Properties for Related Element": bool(False),
        "Check Omniclass Depth": bool(False),
        "Check Property Exists": bool(False),
        "Check Property Value Equals": bool(False),
        "Check Property Value Equals Any of Multiple Values": bool(False),
        "Check Property Value Exists": bool(False),
        "Check Property Value Exists Against Quantity": bool(False),
        "Check Property Value Is In Range": bool(False),
        "Check Property Value is Unique": bool(False),
        "Check Property Value Matches Pattern": bool(False),
        "Ensure Property Missing": bool(False),
        "Ensure Property Value Missing": bool(False),
        "Ensure Property Value Not Equal": bool(False),
        "Check Multiple Properties - 2 args": bool(False)   
    },
    "Parameters": {
        "Name":{"stage":"",
                "value":"",
                "intyp":"",
                "patern":""
                },
        "CreatedBy": {"stage":"",
                      "value":"",
                      "intyp":"",
                      "patern":""
                      },
        "CreatedOn":{"stage":"",
                     "value":"",
                     "intyp":"",
                     "patern":""
                      },
        "Category": {"stage":"",
                     "value":"",
                     "intyp":"",
                     "patern":""
                      },
        "ComponentNames":{"stage":"",
                          "value":"",
                          "intyp":"",
                          "patern":""
                      },
        "Extsystem":{"stage":"",
                     "value":"",
                     "intyp":"",
                     "patern":""
                      },
        "ExtObject":{"stage":"",
                     "value":"",
                     "intyp":"",
                     "patern":""
                      },
        "ExtIdentifier":{"stage":"",
                         "value":"",
                         "intyp":"",
                         "patern":""
                          },
                      },
        "Description":{"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
    }



componentDic ={
    "Varibles": {
        "NumberOfRules": 0,
        "NumberOfParameters": 15
    },
    "ID": {
        "Classification": "",
        "Model": "",
        "Role":"",
        "Description": ""
    },
    "excelSheetLocation": {
        "CellStart":"CM65",
        "CellEnd":"DA74"
    },
    "Rules": {
        "Check Multiple Properties": bool(False) ,
        "Check Multiple Properties - m2": bool(False),
        "Check Multiple Properties - mm": bool(False),
        "Check Multiple Properties for Related Element": bool(False),
        "Check Omniclass Depth": bool(False),
        "Check Property Exists": bool(False),
        "Check Property Value Equals": bool(False),
        "Check Property Value Equals Any of Multiple Values": bool(False),
        "Check Property Value Exists": bool(False),
        "Check Property Value Exists Against Quantity": bool(False),
        "Check Property Value Is In Range": bool(False),
        "Check Property Value is Unique": bool(False),
        "Check Property Value Matches Pattern": bool(False),
        "Ensure Property Missing": bool(False),
        "Ensure Property Value Missing": bool(False),
        "Ensure Property Value Not Equal": bool(False),
        "Check Multiple Properties - 2 args": bool(False)   
    },
   "Parameters": {
        "Name":{"stage":"",
                "value":"",
                "intyp":"",
                "patern":""
                },
        "CreatedBy":{"stage":"",
                     "value":"",
                     "intyp":"",
                     "patern":""
                     },
        "CreatedOn":{"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                      },
        "TypeName ":{"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                      },
        "Space":{"stage":"",
                 "value":"",
                 "intyp":"",
                 "patern":""
                       },
        "Description":{"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                       },
        "Extsystem":{"stage":"",
                     "value":"",
                     "intyp":"",
                     "patern":""
                     },
        "ExtObject": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                     },
        "ExtIdentifier":{"stage":"",
                         "value":"",
                         "intyp":"",
                         "patern":""
                        },
        "SerialNumber": {"stage":"",
                         "value":"",
                         "intyp":"",
                         "patern":""
                        },
        "InstallationDate":{"stage":"",
                            "value":"",
                            "intyp":"",
                            "patern":""
                          },
        "WarrantyStartDate":{"stage":"",
                             "value":"",
                             "intyp":"",
                             "patern":""
                            },
        "TagNumber":{"stage":"",
                     "value":"",
                     "intyp":"",
                     "patern":""
                     },
        "BarCode": {"stage":"",
                    "value":"",
                    "intyp":"",
                    "patern":""
                          },
        "AssetIdentifier": {"stage":"",
                            "value":"",
                            "intyp":"",
                            "patern":""
                          }
    }
}
typeDic ={
    "Varibles": {
        "NumberOfRules": 0,
        "NumberOfParameters": 24
    },
    "ID": {
        "Classification": "",
        "Model": "",
        "Role":"",
        "Description": ""
    },
    "excelSheetLocation": {
        "CellStart":"CM56",
        "CellEnd":"DI56"
    },
    "Rules": {
        "Check Multiple Properties": bool(False) ,
        "Check Multiple Properties - m2": bool(False),
        "Check Multiple Properties - mm": bool(False),
        "Check Multiple Properties for Related Element": bool(False),
        "Check Omniclass Depth": bool(False),
        "Check Property Exists": bool(False),
        "Check Property Value Equals": bool(False),
        "Check Property Value Equals Any of Multiple Values": bool(False),
        "Check Property Value Exists": bool(False),
        "Check Property Value Exists Against Quantity": bool(False),
        "Check Property Value Is In Range": bool(False),
        "Check Property Value is Unique": bool(False),
        "Check Property Value Matches Pattern": bool(False),
        "Ensure Property Missing": bool(False),
        "Ensure Property Value Missing": bool(False),
        "Ensure Property Value Not Equal": bool(False),
        "Check Multiple Properties - 2 args": bool(False)   
    },
    "Parameters": {
        "Name": {"stage":"",
                 "value":"",
                 "intyp":"",
                 "patern":""
                 },
        "CreatedBy":{"stage":"",
                     "value":"",
                     "intyp":"",
                     "patern":""
                     },
        "CreatedOn":{"stage":"",
                     "value":"",
                     "intyp":"",
                     "patern":""
                    },
        "Category": {"stage":"",
                     "value":"",
                     "intyp":"",
                     "patern":""
                     },
        "Description": {"stage":"",
                        "value":"",
                        "intyp":"",
                        "patern":""
                        },
        "AssetType": {"stage":"",
                      "value":"",
                      "intyp":"",
                      "patern":""
                          },
        "Manufacturer": {"stage":"",
                         "value":"",
                         "intyp":"",
                         "patern":""
                          },
        "ModelNumber":{"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "WarrantyGuarantorParts": {"stage":"",
                                   "value":"",
                                   "intyp":"",
                                   "patern":""
                                      },
        "WarrantyDurationParts": {"stage":"",
                                   "value":"",
                                   "intyp":"",
                                   "patern":""
                                      },
        "WarrantyGuarantorLabor": {"stage":"",
                                   "value":"",
                                   "intyp":"",
                                   "patern":""
                                      },
        "WarrantyDurationLabor":{"stage":"",
                                 "value":"",
                                 "intyp":"",
                                 "patern":""
                                      },
        "WarrantyDurationUnit": {"stage":"",
                                 "value":"",
                                 "intyp":"",
                                 "patern":""
                                      },
        "ExtSystem": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "ExtObject":{"stage":"",
                     "value":"",
                     "intyp":"",
                     "patern":""
                          },
        "ExtIdentifier": {"stage":"",
                          "value":"",
                          "intyp":"",
                          "patern":""
                              },
        "ReplacementCost": {"stage":"",
                            "value":"",
                            "intyp":"",
                            "patern":""
                                  },
        "ExpectedLife": {"stage":"",
                         "value":"",
                         "intyp":"",
                         "patern":""
                              },
        "DurationUnit": {"stage":"",
                         "value":"",
                         "intyp":"",
                         "patern":""
                              },
        "WarrantyDescription": {"stage":"",
                                "value":"",
                                "intyp":"",
                                "patern":""
                                  },
        "NominalLength": {"stage":"",
                          "value":"",
                          "intyp":"",
                          "patern":""
                              },
        "NominalWidth": {"stage":"",
                         "value":"",
                         "intyp":"",
                         "patern":""
                              },
        "NominalHeight": {"stage":"",
                          "value":"",
                          "intyp":"",
                          "patern":""
                              }
    }
}
zoneDic ={
    "Varibles": {
        "NumberOfRules": 0,
        "NumberOfParameters": 8
    },
    "ID": {
        "Classification": "",
        "Model": "",
        "Role":"",
        "Description": ""
    },
    "excelSheetLocation": {
         "CellStart":"CM70",
        "CellEnd":"CU72"
    },
    "Rules": {
        "Check Multiple Properties": bool(False) ,
        "Check Multiple Properties - m2": bool(False),
        "Check Multiple Properties - mm": bool(False),
        "Check Multiple Properties for Related Element": bool(False),
        "Check Omniclass Depth": bool(False),
        "Check Property Exists": bool(False),
        "Check Property Value Equals": bool(False),
        "Check Property Value Equals Any of Multiple Values": bool(False),
        "Check Property Value Exists": bool(False),
        "Check Property Value Exists Against Quantity": bool(False),
        "Check Property Value Is In Range": bool(False),
        "Check Property Value is Unique": bool(False),
        "Check Property Value Matches Pattern": bool(False),
        "Ensure Property Missing": bool(False),
        "Ensure Property Value Missing": bool(False),
        "Ensure Property Value Not Equal": bool(False),
        "Check Multiple Properties - 2 args": bool(False)   
    },
    "Parameters": {
        "Name": {"stage":"",
                 "value":"",
                 "intyp":"",
                 "patern":""
                  },
        "CreatedBy":{"stage":"",
                     "value":"",
                     "intyp":"",
                     "patern":""
                      },
        "CreatedOn":{"stage":"",
                     "value":"",
                     "intyp":"",
                     "patern":""
                      },
        "Category":{"stage":"",
                    "value":"",
                    "intyp":"",
                    "patern":""
                      },
        "SpacesName": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Extsystem": {"stage":"",
                      "value":"",
                      "intyp":"",
                      "patern":""
                          },
        "ExtObject": {"stage":"",
                      "value":"",
                      "intyp":"",
                      "patern":""
                          },
        "ExtIdentifier": {"stage":"",
                          "value":"",
                          "intyp":"",
                          "patern":""
                              }
    }
}
spaceDic ={
    "Varibles": {
        "NumberOfRules": 0,
        "NumberOfParameters": 13
    },
    "ID": {
        "Classification": "",
        "Model": "",
        "Role":"",
        "Description": ""
    },
    "excelSheetLocation": {
        "CellStart":"CM58",
        "CellEnd":"CY60"
   },
    "Rules": {
        "Check Multiple Properties": bool(False) ,
        "Check Multiple Properties - m2": bool(False),
        "Check Multiple Properties - mm": bool(False),
        "Check Multiple Properties for Related Element": bool(False),
        "Check Omniclass Depth": bool(False),
        "Check Property Exists": bool(False),
        "Check Property Value Equals": bool(False),
        "Check Property Value Equals Any of Multiple Values": bool(False),
        "Check Property Value Exists": bool(False),
        "Check Property Value Exists Against Quantity": bool(False),
        "Check Property Value Is In Range": bool(False),
        "Check Property Value is Unique": bool(False),
        "Check Property Value Matches Pattern": bool(False),
        "Ensure Property Missing": bool(False),
        "Ensure Property Value Missing": bool(False),
        "Ensure Property Value Not Equal": bool(False),
        "Check Multiple Properties - 2 args": bool(False)   
    },
    "Parameters": {
        "Name":   {"stage":"",
                   "value":"",
                   "intyp":"",
                   "patern":""
                      },
        "CreatedBy": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "CreatedOn": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Category": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "FloorName":{"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Description": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Extsystem":{"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "ExtObject": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "ExtIdentifier": {"stage":"",
                           "value":"",
                           "intyp":"",
                           "patern":""
                              },
        "RoomTag":   {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "UsableHeight":  {"stage":"",
                           "value":"",
                           "intyp":"",
                           "patern":""
                              },
        "GrossArea": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "NetArea":    {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          }
    }
}
floorDic ={
    "Varibles": {
        "NumberOfRules": 0,
        "NumberOfParameters": 10
    },
    "ID": {
        "Classification": "",
        "Model": "",
        "Role":"",
        "Description": ""
    },
    "excelSheetLocation": {
        "CellStart":"CM46",
        "CellEnd":"CV48"
    },
    "Rules": {
        "Check Multiple Properties": bool(False) ,
        "Check Multiple Properties - m2": bool(False),
        "Check Multiple Properties - mm": bool(False),
        "Check Multiple Properties for Related Element": bool(False),
        "Check Omniclass Depth": bool(False),
        "Check Property Exists": bool(False),
        "Check Property Value Equals": bool(False),
        "Check Property Value Equals Any of Multiple Values": bool(False),
        "Check Property Value Exists": bool(False),
        "Check Property Value Exists Against Quantity": bool(False),
        "Check Property Value Is In Range": bool(False),
        "Check Property Value is Unique": bool(False),
        "Check Property Value Matches Pattern": bool(False),
        "Ensure Property Missing": bool(False),
        "Ensure Property Value Missing": bool(False),
        "Ensure Property Value Not Equal": bool(False),
        "Check Multiple Properties - 2 args": bool(False)   
    },
    "Parameters": {
        "Name":      {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "CreatedBy": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "CreatedOn":{"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Category": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Extsystem": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "ExtObject": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "ExtIdentifier": {"stage":"",
                           "value":"",
                           "intyp":"",
                           "patern":""
                              },
        "Description": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Elevation": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Height":     {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          }
    }
}
facilityDic ={
    "Varibles": {
        "NumberOfRules": 0,
        "NumberOfParameters": 22
    },
    "ID": {
        "Classification": "",
        "Model": "",
        "Role":"",
        "Description": ""
    },
    "excelSheetLocation": {
        "CellStart":"CM34",
        "CellEnd":"DH36",  
    },
    "Rules": {
        "Check Multiple Properties": bool(False) ,
        "Check Multiple Properties - m2": bool(False),
        "Check Multiple Properties - mm": bool(False),
        "Check Multiple Properties for Related Element": bool(False),
        "Check Omniclass Depth": bool(False),
        "Check Property Exists": bool(False),
        "Check Property Value Equals": bool(False),
        "Check Property Value Equals Any of Multiple Values": bool(False),
        "Check Property Value Exists": bool(False),
        "Check Property Value Exists Against Quantity": bool(False),
        "Check Property Value Is In Range": bool(False),
        "Check Property Value is Unique": bool(False),
        "Check Property Value Matches Pattern": bool(False),
        "Ensure Property Missing": bool(False),
        "Ensure Property Value Missing": bool(False),
        "Ensure Property Value Not Equal": bool(False),
        "Check Multiple Properties - 2 args": bool(False)   
    },
    "Parameters": {
        "Name":      {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "CreatedBy": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "CreatedOn": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Category": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "ProjectName": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "SiteName": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "LinearUnits": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "AreaUnits": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "VolumeUnits": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "CurrencyUnit": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "AreaMeasurement": {"stage":"",
                           "value":"",
                           "intyp":"",
                           "patern":""
                              },
        "ExternalSystem":{"stage":"",
                           "value":"",
                           "intyp":"",
                           "patern":""
                              },
        "ExternalProjectObject": {"stage":"",
                                   "value":"",
                                   "intyp":"",
                                   "patern":""
                                      },
        "ExternalProjectIdentifier": {"stage":"",
                                       "value":"",
                                       "intyp":"",
                                       "patern":""
                                          },
        "ExternalSiteObject": {"stage":"",
                               "value":"",
                               "intyp":"",
                               "patern":""
                                  },
        "ExternalSiteIdentifier": {"stage":"",
                                   "value":"",
                                   "intyp":"",
                                   "patern":""
                                      },
        "ExternalFacilityObject":{"stage":"",
                                   "value":"",
                                   "intyp":"",
                                   "patern":""
                                      },
        "ExternalFacilityIdentifier": {"stage":"",
                                       "value":"",
                                       "intyp":"",
                                       "patern":""
                                          },
        "Description": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "ProjectDescription": {"stage":"",
                               "value":"",
                               "intyp":"",
                               "patern":""
                                  },
        "SiteDescription":{"stage":"",
                           "value":"",
                           "intyp":"",
                           "patern":""
                              },
        "Phase": {"stage":"",
                   "value":"",
                   "intyp":"",
                   "patern":""
                      }
    }
}
contactDic ={
    "Varibles": {
        "NumberOfRules": 0,
        "NumberOfParameters": 19
    },
    "ID": {
        "Classification": "",
        "Model": "",
        "Role":"",
        "Description": ""
    },
    "excelSheetLocation": {
        "CellStart":"CM22",
        "CellEnd":"DE24"
    },
   "Rules": {
        "Check Multiple Properties": bool(False) ,
        "Check Multiple Properties - m2": bool(False),
        "Check Multiple Properties - mm": bool(False),
        "Check Multiple Properties for Related Element": bool(False),
        "Check Omniclass Depth": bool(False),
        "Check Property Exists": bool(False),
        "Check Property Value Equals": bool(False),
        "Check Property Value Equals Any of Multiple Values": bool(False),
        "Check Property Value Exists": bool(False),
        "Check Property Value Exists Against Quantity": bool(False),
        "Check Property Value Is In Range": bool(False),
        "Check Property Value is Unique": bool(False),
        "Check Property Value Matches Pattern": bool(False),
        "Ensure Property Missing": bool(False),
        "Ensure Property Value Missing": bool(False),
        "Ensure Property Value Not Equal": bool(False),
        "Check Multiple Properties - 2 args": bool(False)   
    },
    "Parameters": {
        "Email":     {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "CreatedBy": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "CreatedOn": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Category": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Company":   {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Phone":     {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "ExtSystem": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "ExtObject": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "ExtIdentifier": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Department": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "OrganizationCode": {"stage":"",
                           "value":"",
                           "intyp":"",
                           "patern":""
                              },
        "GivenName":{"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "FamilyName": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Street":    {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "PostalBox": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Town":       {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "StateRegion":{"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "PostalCode": {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Country":    {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          }
    }
}
instructionDic ={
    "Varibles": {
        "NumberOfRules": 0,
        "NumberOfParameters": 5
    },
    "ID": {
        "Classification": "",
        "Model": "",
        "Role":"",
        "Description": ""
    },
    "excelSheetLocation": {
        "CellStart":"CM10",
        "CellEnd":"CQ12"
    },
    "Rules": {
        "Check Multiple Properties": bool(False) ,
        "Check Multiple Properties - m2": bool(False),
        "Check Multiple Properties - mm": bool(False),
        "Check Multiple Properties for Related Element": bool(False),
        "Check Omniclass Depth": bool(False),
        "Check Property Exists": bool(False),
        "Check Property Value Equals": bool(False),
        "Check Property Value Equals Any of Multiple Values": bool(False),
        "Check Property Value Exists": bool(False),
        "Check Property Value Exists Against Quantity": bool(False),
        "Check Property Value Is In Range": bool(False),
        "Check Property Value is Unique": bool(False),
        "Check Property Value Matches Pattern": bool(False),
        "Ensure Property Missing": bool(False),
        "Ensure Property Value Missing": bool(False),
        "Ensure Property Value Not Equal": bool(False),
        "Check Multiple Properties - 2 args": bool(False)   
    },
    "Parameters": {
        "Title":      {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Version":    {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Release":    {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Status":     {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          },
        "Region":     {"stage":"",
                       "value":"",
                       "intyp":"",
                       "patern":""
                          }  
    }
}








tableDic = {
    "system": {
        "CellStart":"CM74",
        "CellEnd":"CU74",
        "Name": "",
        "CreatedBy": "",
        "CreatedOn": "",
        "Category": "",
        "ComponentNames": "",
        "Extsystem": "",
        "ExtObject": "",
        "ExtIdentifier": "",
        "Description": ""
    },
    "Component": {
        "CellStart":"CM65",
        "CellEnd":"DA74",
        "Name": "",
        "CreatedBy": "",
        "CreatedOn": "",
        "TypeName ": "",
        "Space": "",
        "Description": "",
        "Extsystem": "",
        "ExtObject": "",
        "ExtIdentifier": "",
        "SerialNumber": "",
        "InstallationDate": "",
        "WarrantyStartDate": "",
        "TagNumber": "",
        "BarCode": "",
        "AssetIdentifier": ""
    },
    "Type": {
        "CellStart":"CM56",
        "CellEnd":"DI56",
        "Name": "",
        "CreatedBy": "",
        "CreatedOn": "",
        "Category  ": "",
        "Description": "",
        "AssetType": "",
        "Manufacturer": "",
        "ModelNumber": "",
        "WarrantyGuarantorParts": "",
        "WarrantyDurationParts ": "",
        "WarrantyGuarantorLabor ": "",
        "WarrantyDurationLabor": "",
        "WarrantyDurationUnit": "",
        "ExtSystem": "",
        "ExtObject": "",
        "ExtIdentifier ": "",
        "ReplacementCost ": "",
        "ExpectedLife": "",
        "DurationUnit": "",
        "WarrantyDescription ": "",
        "NominalLength": "",
        "NominalWidth": "",
        "NominalHeight": ""   
    },
    "Zone": {
        "CellStart":"CM47",
        "CellEnd":"CU47",
        "Name": "",
        "CreatedBy": "",
        "CreatedOn": "",
        "Category": "",
        "SpacesName": "",
        "Extsystem": "",
        "ExtObject": "",
        "ExtIdentifier": ""  
    },
    "Space": {
        "CellStart":"CM38",
        "CellEnd":"CY38",
        "Name": "",
        "CreatedBy": "",
        "CreatedOn": "",
        "Category": "",
        "FloorName": "",
        "Description": "",
        "Extsystem": "",
        "ExtObject": "",
        "ExtIdentifier": "",
        "RoomTag": "",
        "UsableHeight": "",
        "GrossArea": "",
        "NetArea": ""
    },
    "Floor": {
        "CellStart":"CM29",
        "CellEnd":"CV29",
        "Name": "",
        "CreatedBy": "",
        "CreatedOn": "",
        "Category": "",
        "Extsystem": "",
        "ExtObject": "",
        "ExtIdentifier": "",
        "Description": "",
        "Elevation": "",
        "Height": ""
    },
    "Facility": {
        "CellStart":"CM20",
        "CellEnd":"DH20",
        "Name": "",
        "CreatedBy": "",
        "CreatedOn": "",
        "Category": "",
        "ProjectName": "",
        "SiteName": "",
        "LinearUnits": "",
        "AreaUnits": "",
        "VolumeUnits": "",
        "CurrencyUnit": "",
        "AreaMeasurement": "",
        "ExternalSystem": "",
        "ExternalProjectObject": "",
        "ExternalProjectIdentifier": "",
        "ExternalSiteObject": "",
        "ExternalSiteIdentifier": "",
        "ExternalFacilityObject": "",
        "ExternalFacilityIdentifier": "",
        "Description": "",
        "ProjectDescription": "",
        "SiteDescription": "",
        "Phase": ""   
    },
    "Contact": {
        "CellStart":"CM11",
        "CellEnd":"DE11",
        "Email": "",
        "CreatedBy": "",
        "CreatedOn": "",
        "Category": "",
        "Company": "",
        "Phone": "",
        "ExtSystem": "",
        "ExtObject": "",
        "ExtIdentifier": "",
        "Department": "",
        "OrganizationCode": "",
        "GivenName": "",
        "FamilyName": "",
        "Street": "",
        "PostalBox": "",
        "Town": "",
        "StateRegion": "",
        "PostalCode": "",
        "Country": "" 
    },
    "Instruction": {
        "CellStart":"CM2",
        "CellEnd":"CQ2",
        "Title": "",
        "Version": "",
        "Release": "",
        "Status": "",
        "Region": ""   
    }
    

}


