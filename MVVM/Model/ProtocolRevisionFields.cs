using System.Collections.ObjectModel;
using System.Globalization;

namespace protocolPlus.MVVM.Model
{
    public class Tool : IEquatable<Tool>
    {
        public string Identifier { get; set; }
        public string Name { get; set; }
        public string NameTag { get; set; } = "<tool.name>";
        public string Type { get; set; }
        public string TypeTag { get; set; } = "<tool.type>";
        public string AssurancNum { get; set; }
        public string AssuranceNumTag { get; set; } = "<tool.assurance_num>";
        public string VerificationNum { get; set; }
        public string VerificationNumTag { get; set; } = "<tool.verification_num>";
        public string VerificationDate { get; set; }
        public string VerificationDateTag { get; set; } = "<tool.verification_date>";
        public string Ready { get => GetReadiness(); }
        public string ReadyTag { get; set; } = "<tool.ready>";

        public string GetReadiness()
        {
            var readiness = "годен";

            var currentDateTime = DateTime.Now;
            var parsedDate = DateTime.ParseExact(VerificationDate, "dd-MM-yyyy", CultureInfo.InvariantCulture);

            if (parsedDate < currentDateTime)
                readiness = "не годен";

            return readiness;
        }

        public override bool Equals(object obj)
        {
            if (obj is Tool otherTool)
            {
                return Identifier == otherTool.Identifier;
            }

            return false;
        }

        public bool Equals(Tool other)
        {
            return Identifier == other.Identifier;
        }

        public override int GetHashCode()
        {
            return Identifier.GetHashCode();
        }
    }
    public class DropDownItem
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public ObservableCollection<DataItem> Data { get; set; } = [];
    }
    public class DataItem
    {
        public string MachineType { get; set; }
        public string MachineTypeTag { get; } = "<>";
        public string MachineName { get; set; }
        public string MachineNameTag { get; } = "<machine.name>";
        public string MachineAssuranceNum { get; set; }
        public string MachineAssuranceNumTag { get; } = "<machine.assurance_num>";
        public string MachinePower { get; set; }
        public string MachinePowerTag { get; } = "<machine.power>";
        public string MachineVoltageSt { get; set; }
        public string MachineVoltageStTag { get; } = "<machine.voltage_st>";
        public string MachineCurrentSt { get; set; }
        public string MachineCurrentStTag { get; } = "<machine.current_st>";
        public string MachineFrequency { get; set; }
        public string MachineFrequencyTag { get; } = "<machine.frequency>";
        public string MachineRpm { get; set; }
        public string MachineRpmTag { get; } = "<machine.rpm>";
        public string MachineCosinus { get; set; }
        public string MachineCosinusTag { get; } = "<machine.cosinus>";
        public string MachineEfficency { get; set; }
        public string MachineEfficencyTag { get; } = "<machine.efficency>";
        public string MachineCurrentExc { get; set; }
        public string MachineCurrentExcTag { get; } = "<machine.current.exec>";
        public string MachineVoltageExc { get; set; }
        public string MachineVoltageExcTag { get; } = "<machine.volatage_exc>";
        public string MachineRotatio { get; set; }
        public string MachineRotatioTag { get; } = "<machine.rotation>";
    }

    class ProtocolRevisionGroups
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public List<ProtocolRevisionFields> FieldsInGroup { get; set; }
    }

    public class ProtocolRevisionFields
    {
        public int ResultId { get; set; }
        public int GroupId { get; set; }
        public string ProtocolFieldQuestion { get; set; }
        public string ProtocolFieldAnswer { get; set; }
        public string ProtocolFieldTag { get; set; }
        public string ProtocolFieldUnits { get; set; }
        public int IsChecked { get; set; }
    }

    class ProtocolRevisionResultFields
    {
        public int Id { get; set; }
        public string Tag { get; set; }
        public int CheckType { get; set; }
        public double PermissibleValue { get; set; }
        public double PermissibleValue2 { get; set; } = 60;
        public List<double> InputValues { get; set; } = [];
        public string Result { get; set; }

        public string GetResult()
        {

            Result = "соответствует";


            int i = 0;
            if (CheckType == 1) // не менее
            {
                while (i < InputValues.Count() && Result == "соответствует")
                {
                    if (InputValues[i] < PermissibleValue)
                        Result = "не соответствует";
                    i++;
                }
            }
            else if (CheckType == 2) // не более
            {
                while (i < InputValues.Count() && Result == "соответствует")
                {
                    if (InputValues[i] > PermissibleValue)
                        Result = "не соответствует";
                    i++;
                }
            }
            else if (CheckType == 3) // отклонение не более между результатами
            {
                while (i < InputValues.Count() && Result == "соответствует")
                {
                    int j = 0;
                    while (j < InputValues.Count() && Result == "соответствует")
                    {
                        double tmp;
                        if (InputValues[j] < InputValues[i])
                            tmp = 100 - ((InputValues[j] / InputValues[i]) * 100);
                        else
                            tmp = 100 - ((InputValues[i] / InputValues[j]) * 100);

                        if (tmp > PermissibleValue)
                            Result = "не соответствует";
                        j++;
                    }
                    i++;
                }
            }
            else if (CheckType == 4) // отклонение не менее между результатами
            {
                while (i < InputValues.Count() && Result == "соответствует")
                {
                    int j = 0;
                    while (j < InputValues.Count() && Result == "соответствует")
                    {
                        double tmp;
                        if (InputValues[j] < InputValues[i])
                            tmp = 100 - ((InputValues[j] / InputValues[i]) * 100);
                        else
                            tmp = 100 - ((InputValues[i] / InputValues[j]) * 100);

                        if (tmp < PermissibleValue)
                            Result = "не соответствует";
                        j++;
                    }
                    i++;
                }
            }
            else if (CheckType == 5) // равно
            {
                while (i < InputValues.Count() && Result == "соответствует")
                {
                    if (InputValues[i] != PermissibleValue)
                        Result = "не соответствует";
                    i++;
                }
            }
            else if (CheckType == 6) // диапазон
            {
                while (i < InputValues.Count() && Result == "соответствует")
                {
                    if ((InputValues[i] < PermissibleValue) || (InputValues[i] > PermissibleValue2))
                        Result = "не соответствует";
                    i++;
                }
            }

            return Result;
        }
    }
}
