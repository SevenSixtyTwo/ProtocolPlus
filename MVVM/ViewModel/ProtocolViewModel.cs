using protocolPlus.MVVM.Model;
using System.Collections.ObjectModel;
using protocolPlus.Core;
using System.IO;
using System.Data.SQLite;
using System.Windows.Input;
using System.ComponentModel;

using static protocolPlus.Core.DatabaseUtils;
using static WordUtils;
using System.Data;

namespace protocolPlus.MVVM.ViewModel
{
    class ProtocolViewModel : ObservableObject
    {
        public ObservableCollection<ProtocolRevisionGroups> ProtocolGroups { get; set; }
        public ObservableCollection<Tool> Tools { get; set; } 
        public ObservableCollection<Tool> AvailableTools { get; set; } 
        public ObservableCollection<DropDownItem> DropDownListItems { get; set; }
        public ObservableCollection<DataItem> DataGridItems { get; set; }
        public ObservableCollection<ProtocolRevisionFields> ProtocolFields { get; set; }
        public Dictionary<int, ProtocolRevisionResultFields> ResultFields { get; set; }

        public string templateFilePath = $@"template-std.docx";
        public string newFileName = "протокол ПСИ СТД";
        public string newFilePath = Environment.CurrentDirectory+@"\протоколы";

        public SQLiteConnection dbConnection = CreateConnection();
        public RelayCommand SaveProtocolCommand { get; set; }
        public Tool SelectedTool { get; set; }

        public ICommand AddToolCommand { get; private set; }
        public ICommand DeleteToolCommand { get; private set; }

        private DropDownItem _selectedItem;
        public DropDownItem SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value;
                UpdateDataGridItems(); // Update data grid when selection changes
                OnPropertyChanged(nameof(SelectedItem));
            }
        }
        private void UpdateDataGridItems()
        {
            DataGridItems.Clear();
            if (SelectedItem != null)
            {
                foreach (var item in SelectedItem.Data)
                {
                    DataGridItems.Add(item);
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public int GetSelectedItemId()
        {
            if (SelectedItem != null)
                return SelectedItem.Id;
            else
                return -1;
        }



        private void AddTool()
        {
            if (SelectedTool != null)
            {
                Tools.Add(new Tool() // Create a copy of the selected tool
                {
                    Identifier = SelectedTool.Identifier,
                    Name = SelectedTool.Name,
                    Type = SelectedTool.Type,
                    AssurancNum = SelectedTool.AssurancNum,
                    VerificationNum = SelectedTool.VerificationNum,
                    VerificationDate = SelectedTool.VerificationDate

                });
                OnPropertyChanged(nameof(Tools));
            }
        }

        private void DeleteTool()
        {
            if (SelectedTool != null)
            {
                //Tool toolToRemove = new Tool()
                //{
                //    Identifier = SelectedTool.Identifier,
                //    Name = SelectedTool.Name,
                //    Cost = SelectedTool.Cost
                //};
                Tools.Remove(SelectedTool);
                SelectedTool = null; // Reset SelectedTool after deletion
                OnPropertyChanged(nameof(Tools));
                OnPropertyChanged(nameof(SelectedTool)); // Notify UI about selection change
            }
        }

        private bool CanDeleteTool()
        {
            return Tools.Count > 0;
        }


        //private string _value;

        //public string Value
        //{
        //    get { return _value; }
        //    set 
        //    {
        //        _value = value; 
        //        OnPropetyChanged();
        //    }
        //}

        public ProtocolViewModel() 
        {
            Tools = [];
            AvailableTools = [];

            ProtocolFields = [];
            ResultFields = [];
            ProtocolGroups = [];

            DropDownListItems = [];
            DataGridItems = [];

            SelectedItem = DropDownListItems.FirstOrDefault();


            dbConnection.Open();

            var cmdTools = dbConnection.CreateCommand();
            cmdTools.CommandText = "SELECT id, name, type, assurance_num, verification_num, verification_date from tool;";

            var readerTools = cmdTools.ExecuteReader();
            while (readerTools.Read())
            {
                AvailableTools.Add(new Tool() { 
                    Identifier = readerTools.GetInt32(0).ToString(), 
                    Name = readerTools.GetString(1),
                    Type = readerTools.GetString(2), 
                    AssurancNum = readerTools.GetString(3), 
                    VerificationNum = readerTools.GetString(4), 
                    VerificationDate = readerTools.GetString(5)
                });
            }

            var cmdMavhine = dbConnection.CreateCommand();
            cmdMavhine.CommandText = "SELECT type, name, assurance_num, power, voltage_st, current_st, frequency, rpm, cosinus, efficency, current_exc, voltage_exc, rotation FROM machine";

            var readerMachine = cmdMavhine.ExecuteReader();

            int machineId = 0;
            while (readerMachine.Read())
            {
                DropDownListItems.Add(new DropDownItem
                {
                    Id = machineId,
                    Name = readerMachine.GetString(1),
                    Data =
                    [
                        new DataItem()
                        {
                            MachineType = readerMachine.GetString(0),
                            MachineName = readerMachine.GetString(1),
                            MachineAssuranceNum = readerMachine.GetString(2),
                            MachinePower = readerMachine.GetString(3),
                            MachineVoltageSt = readerMachine.GetString(4),
                            MachineCurrentSt = readerMachine.GetString(5),
                            MachineFrequency = readerMachine.GetString(6),
                            MachineRpm = readerMachine.GetString(7),
                            MachineCosinus = readerMachine.GetString(8),
                            MachineEfficency = readerMachine.GetString(9),
                            MachineCurrentExc = readerMachine.GetString(10),
                            MachineVoltageExc = readerMachine.GetString(11),
                            MachineRotatio = readerMachine.GetString(12)
                        }
                    ]
                });
                machineId++;
            }

            var cmdGroup = dbConnection.CreateCommand();
            cmdGroup.CommandText = "SELECT revisionGroup.id, " +
                "revisionGroup.name, " +
                "revisionCell.revisionResult_id, " +
                "revisionCell.name, " +
                "revisionCell.tag, " +
                "revisionCell.is_checked, " +
                "revisionCell.units, " +
                "revisionCell.id " +
                "FROM revisionGroup JOIN revisionCell ON revisionCell.revisionGroup_id = revisionGroup.id";

            int grpIdCheck = -1;
            int tmpId = 0;
            int fieldId = 0;

            string format1 = "{0}, ({1})";
            string format2 = "{0}";

            string frmt;

            var readerGroup = cmdGroup.ExecuteReader();
            while (readerGroup.Read())
            {
                tmpId = 0;
                if (!readerGroup.IsDBNull(2))
                    tmpId = readerGroup.GetInt32(2);

                frmt = format1;
                if (readerGroup.GetString(6) == "0")
                    frmt = format2;

                int grpId = readerGroup.GetInt32(0)-1;

                if (grpIdCheck != grpId)
                {
                    grpIdCheck = grpId;
                    fieldId = 0;

                    ProtocolGroups.Add(new ProtocolRevisionGroups
                    {
                        Name = readerGroup.GetString(1),
                        FieldsInGroup = []
                    });
                }

                ProtocolGroups[grpId].FieldsInGroup.Add(new ProtocolRevisionFields
                {
                    ResultId = tmpId,
                    ProtocolFieldQuestion = String.Format(frmt, readerGroup.GetString(3), readerGroup.GetString(6)),
                    ProtocolFieldTag = readerGroup.GetString(4),
                    IsChecked = readerGroup.GetInt32(5),
                    ProtocolFieldAnswer = "0",
                    GroupId = grpId
                });

                ProtocolFields.Add(ProtocolGroups[grpId].FieldsInGroup[fieldId]);

            }
            readerGroup.Close();
            cmdGroup.Reset();

            //int tmpId = 0;

            //var cmd = dbConnection.CreateCommand();
            //cmd.CommandText = "SELECT revisionResult_id, name, tag, is_checked, units, revisionGroup_id FROM revisionCell";

            //string format1 = "{0}, ({1})";
            //string format2 = "{0}";

            //string frmt;

            //var reader = cmd.ExecuteReader();
            //while (reader.Read())
            //{
            //    tmpId = 0;
            //    if (!reader.IsDBNull(0))
            //        tmpId = reader.GetInt32(0);

            //    frmt = format1;
            //    if (reader.GetString(4) == "0")
            //        frmt = format2;

            //    ProtocolFields.Add(new ProtocolRevisionFields
            //    {
            //        ResultId = tmpId,
            //        ProtocolFieldQuestion = String.Format(frmt, reader.GetString(1), reader.GetString(4)),
            //        ProtocolFieldTag = reader.GetString(2),
            //        IsChecked = reader.GetInt32(3),
            //        ProtocolFieldAnswer = "0",
            //        GroupId = reader.GetInt32(5)
            //    });
            //}
            //reader.Close();
            //cmd.Reset();
            
            var cmd2 = dbConnection.CreateCommand();
            cmd2.CommandText = "SELECT id, tag, permissible_value, check_type FROM revisionResult";
            
            var reader2 = cmd2.ExecuteReader();
            while (reader2.Read())
            {
                ResultFields.Add(reader2.GetInt32(0), 
                    new ProtocolRevisionResultFields
                    {
                        Tag = reader2.GetString(1),
                        PermissibleValue = Convert.ToDouble(reader2.GetString(2)),
                        CheckType = reader2.GetInt32(3)                   
                    }
                );
            } 
            reader2.Close();



            AddToolCommand = new RelayCommand(o =>
            {
                AddTool();
            });
            DeleteToolCommand = new RelayCommand(o =>
            {
                if(CanDeleteTool())
                    DeleteTool();
            });



            SaveProtocolCommand = new RelayCommand(o =>
            {
                string currentDate = DateTime.Now.ToString("dd.MM.yyyy");
                string finalResult = "пригоден";

                Dictionary<string, string> tagsAndValues = [];

                tagsAndValues.Add("<protocol.creation_date>", currentDate);

                var newFullPathFile = CreateNewDocumentFromTemplate(templateFilePath, newFilePath, newFileName);

                foreach (var group in ProtocolGroups)
                {
                    foreach (var field in group.FieldsInGroup)
                    {
                        tagsAndValues.Add(field.ProtocolFieldTag, field.ProtocolFieldAnswer);

                        if (field.ResultId != 0)
                        {
                            double answer = Convert.ToDouble(field.ProtocolFieldAnswer);
                            int tmpID = field.ResultId;

                            ResultFields[tmpID].InputValues.Add(answer);
                        }
                    }
                }

                int selectedMachineId = GetSelectedItemId();

                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineNameTag, DropDownListItems[selectedMachineId].Data[0].MachineName);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineAssuranceNumTag, DropDownListItems[selectedMachineId].Data[0].MachineAssuranceNum);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachinePowerTag, DropDownListItems[selectedMachineId].Data[0].MachinePower);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineVoltageStTag, DropDownListItems[selectedMachineId].Data[0].MachineVoltageSt);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineCurrentStTag, DropDownListItems[selectedMachineId].Data[0].MachineCurrentSt);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineFrequencyTag, DropDownListItems[selectedMachineId].Data[0].MachineFrequency);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineRpmTag, DropDownListItems[selectedMachineId].Data[0].MachineRpm);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineCosinusTag, DropDownListItems[selectedMachineId].Data[0].MachineCosinus);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineEfficencyTag, DropDownListItems[selectedMachineId].Data[0].MachineEfficency);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineCurrentExcTag, DropDownListItems[selectedMachineId].Data[0].MachineCurrentExc);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineVoltageExcTag, DropDownListItems[selectedMachineId].Data[0].MachineVoltageExc);
                tagsAndValues.Add(DropDownListItems[selectedMachineId].Data[0].MachineRotatioTag, DropDownListItems[selectedMachineId].Data[0].MachineRotatio);


                string tmpRez;
                foreach (var result in ResultFields)
                {
                    tmpRez = result.Value.GetResult();

                    if (tmpRez == "не соответствует")
                        finalResult = "не пригоден";
                    tagsAndValues.Add(result.Value.Tag, tmpRez);
                }

                tagsAndValues.Add("<final_result>", finalResult);
                
                foreach(var item in tagsAndValues)
                {
                    ReplaceTagsInFile(newFullPathFile, item.Key, item.Value);
                }
                
                Dictionary<string, string> toolsTagsAndValues = [];

                string tableCaption = "TOOLS_TABLE";
                var rowPattern = GetRowPattern(newFullPathFile, tableCaption);
                foreach (var tool in Tools)
                {
                    toolsTagsAndValues.Add(tool.NameTag, tool.Name);
                    toolsTagsAndValues.Add(tool.TypeTag, tool.Type);
                    toolsTagsAndValues.Add(tool.AssuranceNumTag, tool.AssurancNum);
                    toolsTagsAndValues.Add(tool.VerificationNumTag, tool.VerificationNum);
                    toolsTagsAndValues.Add(tool.VerificationDateTag, tool.VerificationDate);
                    toolsTagsAndValues.Add(tool.ReadyTag, tool.Ready);

                    CreateRowWithPattern(newFullPathFile, tableCaption, rowPattern, toolsTagsAndValues);
                    toolsTagsAndValues.Clear();
                }
                DeleteRowPattern(newFullPathFile, tableCaption, rowPattern);
            });            
        }
    }
}
