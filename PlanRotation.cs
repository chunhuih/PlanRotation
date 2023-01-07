using System;
using System.Linq;
using System.IO;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Windows.Documents;
using System.Windows.Controls;


///////////////////////////////////////
// 6/22/2022:
// I am updating the code for the following functions:
// 1. Create MLC leaf array dynamically so that it can accommodate different number of MLC leaves.
// 2. Add support for Halcyon / Ethos linacs
//////////////////////////////////////

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.1.3")]
[assembly: AssemblyFileVersion("1.0.1.3")]
[assembly: AssemblyInformationalVersion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
[assembly: ESAPIScript(IsWriteable = true)]

namespace PlanRotation
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string logName = System.AppDomain.CurrentDomain.FriendlyName + ".log";
            using (StreamWriter w = File.AppendText(logName))
            {
                w.AutoFlush = true;
                string log = "Start of the app: " + System.AppDomain.CurrentDomain.FriendlyName;
                Log(log, w);
            }
            try
            {
                using (Application app = Application.CreateApplication())
                {
                    Execute(app);
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
            }
            using (StreamWriter w = File.AppendText(logName))
            {
                w.AutoFlush = true;
                string log = "End of the app: " + System.AppDomain.CurrentDomain.FriendlyName;
                Log(log, w);
            }
        }
        public static void Log(string logMessage, TextWriter w)
        {
            // The $ symbol before the quotation mark creates an interpolated string.
            w.Write($"{DateTime.Now.ToLongTimeString()} {DateTime.Now.ToLongDateString()}");
            w.WriteLine($": {logMessage}");
        }
        private static float[,] Rotated(float[,] orig)
        {
            int L1 = 28;
            int L2 = 57;
            var newLeaves = orig.Clone() as float[,];
            for (int i = 0; i < L1; i++)
            {
                if (Math.Abs(Math.Abs(orig[1, L1 - i - 1]) - 140) < 0.5)
                {
                    newLeaves[0, i] = orig[0, L1 - i - 1];
                }
                else
                {
                    newLeaves[0, i] = -orig[1, L1 - i - 1];
                }
                if (Math.Abs(Math.Abs(orig[0, L1 - i - 1]) - 140) < 0.5)
                {
                    newLeaves[1, i] = orig[1, L1 - i - 1];
                }
                else
                {
                    newLeaves[1, i] = -orig[0, L1 - i - 1];
                }
            }
            for (int i = 28; i < L2; i++)
            {
                if (Math.Abs(Math.Abs(orig[1, L2 + 28 - i - 1] - 140)) < 0.5)
                {
                    newLeaves[0, i] = orig[0, L2 + 28 - i - 1];
                }
                else
                {
                    newLeaves[0, i] = -orig[1, L2 + 28 - i - 1];
                }
                if (Math.Abs(Math.Abs(orig[0, L2 + 28 - i - 1]) - 140) < 0.5)
                {
                    newLeaves[1, i] = orig[1, L2 + 28 - i - 1];
                }
                else
                {
                    newLeaves[1, i] = -orig[0, L2 + 28 - i - 1];
                }
            }
            return newLeaves;
        }
        private static float[,] Flipped(float[,] orig)
        {
            int L1 = 28;
            int allPairs = 57;
            var ri = L1 - 1;
            var li = 0;
            var retval = orig.Clone() as float[,];
            //
            // L1
            //
            for (; li<ri; li++, ri--)
            {
                var tmp1 = retval[0, li];
                var tmp2 = retval[1, li];
                retval[0, li] = retval[0, ri];
                retval[1, li] = retval[1, ri];
                retval[0, ri] = tmp1;
                retval[1, ri] = tmp2;
            }
            //
            // L2
            //
            ri = allPairs - 1;
            li = L1;
            for (; li<ri; li++, ri--)
            {
                var tmp1 = retval[0, li];
                var tmp2 = retval[1, li];
                retval[0, li] = retval[0, ri];
                retval[1, li] = retval[1, ri];
                retval[0, ri] = tmp1;
                retval[1, ri] = tmp2;
            }
            return retval;
        }
        private static float[,] FlippedStatic(float[,] orig)
        {
            int L1 = 28;
            var ri = L1 - 1;
            var li = 0;
            var retval = orig.Clone() as float[,];
            for (; li < ri; li++, ri--)
            {
                var tmp1 = retval[0, li];
                var tmp2 = retval[1, li];
                retval[0, li] = retval[0, ri];
                retval[1, li] = retval[1, ri];
                retval[0, ri] = tmp1;
                retval[1, ri] = tmp2;
            }
            return retval;
        }
        static void Execute(Application app)
        {
            string logName = System.AppDomain.CurrentDomain.FriendlyName + ".log";
            Console.Clear();
            Console.WriteLine("=========================================================================\n");
            Console.WriteLine("This is an external beam plan modification application.\n");
            Console.WriteLine("It rotates treatment beams as a result of patient orientation change.\n");
            Console.WriteLine("Please read application instructions for details.\n");
            Console.WriteLine("=========================================================================\n");
            Console.Write("Please enter the patient ID: ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            string mrn = Console.ReadLine();
            Console.ResetColor();
            var patient = app.OpenPatientById(mrn);
            if (patient == null)
            {
                Console.WriteLine("ERROR: This patient ID does not exist.\n");
                Console.WriteLine("Please use a correct patient ID and run this application again.\n");
                return;
            }
            int nCourses = patient.Courses.Count();
            if (nCourses == 0)
            {
                Console.WriteLine("ERROR: This patient does not contain any course.\n");
                Console.WriteLine("Please choose another patient with existing courses and run this application again.\n");
                return;
            }
            int nPlans = 0;
            foreach (Course eachCourse in patient.Courses)
            {
                nPlans += eachCourse.PlanSetups.Count();
            }
            if (nPlans == 0)
            {
                Console.WriteLine("ERROR: This patient does not contain any plan.\n");
                Console.WriteLine("Please choose another patient with existing plans and run this application again.\n");
                return;
            }
            Console.WriteLine($"Found {nPlans} plan(s) in {nCourses} courses for this patient with ID: {mrn}.");
            Console.WriteLine("Please choose the course and plan from the list below:");
            foreach (Course eachCourse in patient.Courses)
            {
                foreach (PlanSetup eachPlan in eachCourse.PlanSetups)
                {
                    Console.WriteLine($"{patient.Id}\t\"{eachCourse.Id}\"\t\"{eachPlan.Id}\"");
                }
            }
            Console.Write("First please enter the course ID for your plan: ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            string courseID = Console.ReadLine();
            Console.ResetColor();
            var courses = patient.Courses.Where(c => c.Id == courseID);
            if (!courses.Any())
            {
                Console.WriteLine("ERROR: The course ID is not found.");
                Console.WriteLine("Please choose a patient with a course and run this application again.");
                return;
            }
            var course = courses.Single();
            Console.Write("Next, please enter the plan name: ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            string planName = Console.ReadLine();
            Console.ResetColor();
            var plans = course.PlanSetups.Where(p => p.Id == planName);
            if (!plans.Any())
            {
                Console.WriteLine("ERROR: This plan ID is not found. Program will exit.");
                Console.WriteLine("Please choose a patient with an external beam plan and run this application again.");
                return;
            }
            var planSetup = plans.Single();  // It will throw an exception if there is not exactly one instance.
            if (planSetup.PlanType != PlanType.ExternalBeam)
            {
                Console.WriteLine($"ERROR: The plan type is: {planSetup.PlanType}");
                Console.WriteLine("Please choose an external beam plan and run this application again.");
                return;
            }
            // Check if beams are defined for this plan.
            if (planSetup.Beams.Count() == 0)
            {
                Console.WriteLine($"ERROR: The plan has no beams defined.");
                Console.WriteLine("Please choose an external beam plan with beams and run this application again.");
                return;
            }
            // print beam information.
            Console.WriteLine("Here is a list of all the beams in this plan:");
            foreach (var beam in planSetup.Beams)
            {
                Console.Write($"Beam ID: \"{beam.Id}\"\t");
                Console.Write($"Beam name: \"{beam.Name}\"\t");
                string fldType = beam.IsSetupField ? "Setup" : "Tx";
                Console.Write($"Beam type: \"{fldType}\"\t");
                Console.Write($"Machine model: \"{beam.TreatmentUnit.MachineModel}\"\t");
                Console.Write($"Machine model name: \"{beam.TreatmentUnit.MachineModelName}\"\t");
                Console.Write($"Technique: \"{beam.Technique}\"\t");
                Console.WriteLine($"MLC type: \"{beam.MLCPlanType}\"");
            }
            // check if treatment beams are defined.
            int numTxBeams = 0;
            foreach (var beam in planSetup.Beams)
            {
                if (beam.IsSetupField == false && beam.IsImagingTreatmentField == false)
                {
                    numTxBeams++;
                }
            }
            if (numTxBeams == 0)
            {
                Console.WriteLine($"ERROR: The plan has no treatment beams.");
                Console.WriteLine("Please choose an External Beam plan with treatment beams and run this application again.");
                return;
            }
            else
            {
                Console.WriteLine($"{numTxBeams} treatment field(s) were found.");
            }
            // Based on print out results of some test plans, the LeafPositions array for MLC leave positions at each control point 
            // is stored in a two-dimensional array.
            // =================== For the SX2 MLC on a Halcyon / Ethos, the array dimension is [2, 57].
            // The first row of this array [0, 0 --- 56] contains leaf positions for leaves on the X1 bank.
            // The first row of this array [1, 0 --- 56] contains leaf positions for leaves on the X2 bank.
            // The first 28 leaf pairs are ordered from the Y1 side to the Y2 side.
            // The remaining 29 leaf pairs are again ordered from the Y1 side to the Y2 side, with a 5-mm position difference in the 
            // Y direction relative to the first 28 leaf pairs.
            // With a Halcyon/Ethos linac, the collimator is fixed. X1 = 14, X2 = 14, Y1 = -14, Y2 = 14
            // The collimator angle can only be in [270, 359.9] and [0, 90].
            // 
            // =================== For the Millenium 120-leaf MLC, the array dimension is [2, 60].
            // The first row of this array [0, 0 --- 59] contains leaf positions for leaves on the X1 bank.
            // The first row of this array [1, 0 --- 59] contains leaf positions for leaves on the X2 bank.
            // The second index of the leaf array starts with 0 for the most peripheral leaf on the Y1 side, and it ends with 59 
            // for the most peripheral leaf on the Y2 side.
            // This indexing order of the leaves applies to both banks.
            // The jaw and MLC positions are stored in unit of millimeter.
            // Positions for the X jaws and MLC leaves are negative if they are on the X1 side from the mid plane, and are positive
            // they are on the X2 side from the mid plan.
            // Positions for the Y jaws are negative if they are on the Y1 side from the mid plane, and are positive
            // they are on the Y2 side from the mid plan.
            // Based on tests on some plans in Eclipse, here are control point settings based on beam types:
            // For static beams without MLC defined, there are two control points. Each one had control point index = -1.
            // For field-in-field beams, the control point index increases from 0 to n.
            // 
            // When the patient changes from a head-in supine position to a feet-in supine position, the following will be changed
            // for each control point of a beam:
            // 1. Gantry angle of the beam.
            //    According to ESAPI documentation, The GantryAngle property of the ControlPointParameters class cannot be edited.
            //    It only has a getter, not a setter. 
            //    To modify the gantry angle of an existing beam, remove and create a new beam.
            // 2. Collimator angle.
            // 3. Collimator jaws (X1, X2, Y1, Y2).
            // 4. MLC leaf positions at each control point.
            //
            // As far as I can tell, there is no ESAPI method to change patient orientation.
            // There is one property: PlanSetup.TreatmentOrientation, but it does not have a setter.
            // So, this application assumes that the treatment orientation is manually updated by the user before running this application.
            // 
            Console.WriteLine("============================================================================================\n");
            Console.WriteLine("We are going to update the beams in this treatment plan to match patient orientation change.");
            Console.WriteLine("We assume that this treatment plan already has the patient orientation changed.");
            Console.Write($"The orientation of this treatment plan is ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"{planSetup.TreatmentOrientation}.");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("Are you sure that you want to change the beams in this plan? (Y/N): ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            string response = Console.ReadLine();
            Console.ResetColor();
            if (response.ToLower() != "y")
            {
                Console.WriteLine("Application exits. Please choose a suitable plan for this application");
                return;
            }
            ExternalPlanSetup extPlanSetup = planSetup as ExternalPlanSetup;
            Console.Write("Do you want to make a backup copy of this plan first? (Y/N): ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            response = Console.ReadLine();
            Console.ResetColor();
            patient.BeginModifications();
            if (response.ToLower() == "y")
            {
                PlanSetup backupPlan = course.CopyPlanSetup(extPlanSetup);
                for (int i = 1; i <= 99; i++)
                {
                    if (course.PlanSetups.Where(p => p.Id == $"backup{i}").Any() == false)
                    {
                        backupPlan.Id = $"backup{i}";
                        break;
                    }
                }
                Console.WriteLine($"The plan was copied with ID: {backupPlan.Id}");
                app.SaveModifications();
            }
            bool isStaticBeamPlan = false;
            bool isDynamicBeamPlan = false;
            bool isArcBeamPlan = false;
            bool isSX2MLC = false;
            bool isHalcyon = false;
            foreach (var beam in extPlanSetup.Beams)
            {
                if (beam.IsSetupField == false && beam.IsImagingTreatmentField == false)
                {
                    if (beam.Technique.ToString() == "STATIC")
                    {
                        // Here we assume that if one treatment beam is a static beam, then all the beams
                        // in this treatment plan must be static beams.
                        isStaticBeamPlan = true;
                    }
                    if (beam.Technique.ToString().ToLower().Contains("arc"))
                    {
                        // Here we assume that if one treatment beam is a VMAT or arc beam, then all the beams
                        // in this treatment plan must be VMAT or arc beams.
                        isArcBeamPlan = true;
                    }
                    if (beam.MLC.Model == "SX2")
                    {
                        // Here we assume that it is a Halcyon or Ethos linac, because the "SX2" MLC only
                        // appears on such linacs.
                        isSX2MLC = true;
                    }
                    if (beam.TreatmentUnit.MachineModel == "RDS")
                    {
                        // Here we assume that it is a Halcyon or Ethos linac
                        isHalcyon = true;
                    }
                }
            }
            if (isHalcyon && isArcBeamPlan)  // this plan has arc beams on a Halcyon or Ethos
            {
                ExternalBeamMachineParameters[] machineParameters = new ExternalBeamMachineParameters[numTxBeams];
                string[] machineID = new string[numTxBeams];
                string[] beamId = new string[numTxBeams];
                string[] newBeamId = new string[numTxBeams];
                string[] beamName = new string[numTxBeams];
                int[] doseRate = new int[numTxBeams];
                string[] energyModeId = new string[numTxBeams];
                string[] primaryFluenceModeId = new string[numTxBeams];
                string[] technique = new string[numTxBeams];
                MLCPlanType[] mlcType = new MLCPlanType[numTxBeams];
                double[] collimatorAngle = new double[numTxBeams];
                double[] gantryStartAngle = new double[numTxBeams];
                double[] gantryStopAngle = new double[numTxBeams];
                GantryDirection[] gantryDirection = new GantryDirection[numTxBeams];
                int[] numControlPoints = new int[numTxBeams];
                double[] patientSupportAngle = new double[numTxBeams];
                VVector[] isocenterPosition = new VVector[numTxBeams];
                VRect<double>[] jawPositions = new VRect<double>[numTxBeams];
                double[] weightFactor = new double[numTxBeams];
                float[][,] leafPositions = new float[numTxBeams][,];
                var cpsList = new ArrayList();
                int indexTxBeams = 0;
                foreach (Beam beam in extPlanSetup.Beams)
                {
                    if (beam.IsSetupField || beam.IsImagingTreatmentField)
                    {
                        continue;
                    }
                    using (StreamWriter w = File.AppendText(logName))  // record some log data
                    {
                        w.AutoFlush = true;
                        string log = $"Analyzing beam: {beam.Id}: \"{beam.Name}\". Technique: {beam.Technique.Id}. ";
                        log += $"It has {beam.ControlPoints.Count} control points.";
                        Log(log, w);
                        foreach (var cps in beam.ControlPoints)
                        {
                            log = $"Control Point {cps.Index}. Gantry = {cps.GantryAngle}; Collimator = {cps.CollimatorAngle}; ";
                            log += $"X1 Jaw = {cps.JawPositions.X1} mm; ";
                            log += $"X2 Jaw = {cps.JawPositions.X2} mm; ";
                            log += $"Y1 Jaw = {cps.JawPositions.Y1} mm; ";
                            log += $"Y2 Jaw = {cps.JawPositions.Y2} mm; ";
                            log += $"Meterweight = {cps.MetersetWeight}\n";
                            float[,] leaves = cps.LeafPositions;
                            log += $"{leaves.Length} leaf positions are found. ";
                            log += $"Leaf position array rank: {leaves.Rank}. ";
                            for (int i = 0; i < leaves.Rank; i++)
                            {
                                log += $"Array length at dimension {i}: {leaves.GetLength(i)}";
                            }
                            Log(log, w);
                            log = "";
                            for (int i = 0; i < leaves.GetLength(0); i++)
                            {
                                log += $"[{i}, x]: ";
                                for (int j = 0; j < leaves.GetLength(1); j++)
                                {
                                    log += $"{leaves[i, j]}, ";
                                }
                            }
                            Log(log, w);
                        }
                    }
                    ExternalBeamTreatmentUnit machine = beam.TreatmentUnit;
                    machineID[indexTxBeams] = machine.Id;
                    beamId[indexTxBeams] = beam.Id;
                    beamName[indexTxBeams] = beam.Name;
                    doseRate[indexTxBeams] = beam.DoseRate;
                    energyModeId[indexTxBeams] = beam.EnergyModeDisplayName;
                    technique[indexTxBeams] = beam.Technique.ToString();
                    if (energyModeId[indexTxBeams].Contains("-") && energyModeId[indexTxBeams].Split('-')[1] == "FFF")
                    {
                        primaryFluenceModeId[indexTxBeams] = energyModeId[indexTxBeams].Split('-')[1];
                        energyModeId[indexTxBeams] = energyModeId[indexTxBeams].Split('-')[0];
                    }
                    else
                    {
                        primaryFluenceModeId[indexTxBeams] = null;
                    }
                    mlcType[indexTxBeams] = beam.MLCPlanType;
                    weightFactor[indexTxBeams] = beam.WeightFactor;
                    machineParameters[indexTxBeams] = new ExternalBeamMachineParameters(machineID[indexTxBeams], energyModeId[indexTxBeams],
                        doseRate[indexTxBeams], technique[indexTxBeams], primaryFluenceModeId[indexTxBeams]);
                    machineParameters[indexTxBeams].MLCId = @"SX2 MLC";
                    collimatorAngle[indexTxBeams] = beam.ControlPoints[0].CollimatorAngle;
                    patientSupportAngle[indexTxBeams] = beam.ControlPoints[0].PatientSupportAngle;
                    numControlPoints[indexTxBeams] = beam.ControlPoints.Count;
                    isocenterPosition[indexTxBeams] = beam.IsocenterPosition;
                    jawPositions[indexTxBeams] = beam.ControlPoints[0].JawPositions;
                    leafPositions[indexTxBeams] = beam.ControlPoints[0].LeafPositions;
                    gantryStartAngle[indexTxBeams] = beam.ControlPoints.First().GantryAngle;
                    gantryStopAngle[indexTxBeams] = beam.ControlPoints.Last().GantryAngle;
                    gantryDirection[indexTxBeams] = beam.GantryDirection;
                    // Next we rotate the beam.
                    // First rotate the gantry for the beam.
                    if (gantryStartAngle[indexTxBeams] != 0)
                    {
                        gantryStartAngle[indexTxBeams] = 360 - gantryStartAngle[indexTxBeams];
                    }
                    if (gantryStopAngle[indexTxBeams] != 0)
                    {
                        gantryStopAngle[indexTxBeams] = 360 - gantryStopAngle[indexTxBeams];
                    }
                    if (gantryDirection[indexTxBeams] == GantryDirection.Clockwise)
                    {
                        gantryDirection[indexTxBeams] = GantryDirection.CounterClockwise;
                    }
                    else if (gantryDirection[indexTxBeams] == GantryDirection.CounterClockwise)
                    {
                        gantryDirection[indexTxBeams] = GantryDirection.Clockwise;
                    }
                    VRect<double> currentJaws = jawPositions[indexTxBeams];
                    VRect<double> newJaws = new VRect<double>(-currentJaws.X2, -currentJaws.Y2, -currentJaws.X1, -currentJaws.Y1);
                    jawPositions[indexTxBeams] = newJaws;
                    List<CPModel> cpList = new List<CPModel>();
                    float icp = 0;
                    foreach (var cps in beam.ControlPoints)
                    {
                        float[,] currentLeaves = cps.LeafPositions;
                        float[,] newLeaves = Rotated(cps.LeafPositions);
                        if (icp < 2)
                        {
                            for (int i = 0; i < 2; i++)
                            {
                                Console.Write($"i={i}\nL1: ");
                                for (int j = 0; j < 57; j++)
                                {
                                    if (j == 28)
                                    {
                                        Console.Write("\nL2: ");
                                    }
                                    Console.Write($"{cps.LeafPositions[i, j]} ");
                                }
                                Console.WriteLine("");
                            }
                        }
                        if (icp < 2)
                        {
                            Console.WriteLine("Updated leaves: ");
                            for (int i = 0; i < 2; i++)
                            {
                                Console.Write($"i={i}\nL1: ");
                                for (int j = 0; j < 57; j++)
                                {
                                    if(j == 28)
                                    {
                                        Console.Write("\nL2: ");
                                    }
                                    Console.Write($"{newLeaves[i, j]} ");
                                }
                                Console.WriteLine("");
                            }
                        }
                        cpList.Add(new CPModel
                        {
                            GantryAngle = cps.GantryAngle,
                            JawPositions = new VRect<double>(-cps.JawPositions.X2, -cps.JawPositions.Y2, -cps.JawPositions.X1, -cps.JawPositions.Y1),
                            MetersetWeight = cps.MetersetWeight,
                            CollimatorAngle = cps.CollimatorAngle,
                            MLCPositions = newLeaves
                        });
                        icp++;
                    }
                    cpsList.Add(cpList);
                    indexTxBeams++;
                }
                // create new beams with rotated geometry
                for (int idxTxBeam = 0; idxTxBeam < numTxBeams; idxTxBeam++)
                {
                    using (StreamWriter w = File.AppendText(logName))
                    {
                        w.AutoFlush = true;
                        string log = $"Creating a new rotated beam for beam: \"{beamId[idxTxBeam]}\"";
                        Log(log, w);
                    }
                    // ESAPI manual shows that the number of meterset weight items define the number of created control points.
                    List<CPModel> cpList = (List<CPModel>)cpsList[idxTxBeam];
                    var metersetWeights = from cp in cpList select cp.MetersetWeight;
                    Beam newBeam = extPlanSetup.AddVMATBeamForFixedJaws(machineParameters[idxTxBeam], metersetWeights, collimatorAngle[idxTxBeam],
                        gantryStartAngle[idxTxBeam], gantryStopAngle[idxTxBeam], gantryDirection[idxTxBeam],
                        patientSupportAngle[idxTxBeam], isocenterPosition[idxTxBeam]);
                    BeamParameters bParam = newBeam.GetEditableParameters();
                    for (int indCP = 0; indCP < cpList.Count; indCP++)
                    {
                        /*
                        if (indCP < 2)
                        {
                            for (int i = 0; i < 2; i++)
                            {
                                Console.Write($"i={i} ");
                                for (int j = 0; j < 57; j++)
                                {
                                    Console.Write($"{cpList[indCP].MLCPositions[i, j]} ");
                                }
                                Console.WriteLine("");
                            }
                        }
                        */
//                        bParam.ControlPoints.ElementAt(indCP).JawPositions = cpList[indCP].JawPositions;
                        bParam.ControlPoints.ElementAt(indCP).LeafPositions = cpList[indCP].MLCPositions;
                        try
                        {
                            newBeam.ApplyParameters(bParam);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Exception: \n" + ex.ToString());
                        }
                    }
                    //                    bParam.WeightFactor = weightFactor[idxTxBeam];
                    newBeam.Name = beamName[idxTxBeam];
                    newBeamId[idxTxBeam] = newBeam.Id;
                }
                // Remove the old treatment beams.
                foreach (string eachTxBeamId in beamId)
                {
                    try
                    {
                        extPlanSetup.RemoveBeam(extPlanSetup.Beams.Where(b => b.Id == eachTxBeamId).Single());
                    }
                    catch (Exception e)
                    {
                        string log = e.ToString();
                        using (StreamWriter w = File.AppendText(logName))
                        {
                            w.AutoFlush = true;
                            Log(log, w);
                        }
                        Console.Error.WriteLine($"Failed to remove one existing beam: \"{eachTxBeamId}\". Please remove it manually.");
                    }
                }
                // Calculate plan dose with the new beams
                try
                {
                    // first create a <beamID, beam meterset> list as preset values.
                    List<KeyValuePair<string, MetersetValue>> presetValues = new List<KeyValuePair<string, MetersetValue>>();
                    // create new beams with rotated geometry
                    for (int idxTxBeam = 0; idxTxBeam < numTxBeams; idxTxBeam++)
                    {
                        presetValues.Add(new KeyValuePair<string, MetersetValue>(newBeamId[idxTxBeam],
                            extPlanSetup.Beams.FirstOrDefault(b => b.Id == beamId[idxTxBeam]).Meterset));
                    }
                    extPlanSetup.CalculateDoseWithPresetValues(presetValues);
                }
                catch (Exception e)
                {
                    string log = e.ToString();
                    using (StreamWriter w = File.AppendText(logName))
                    {
                        w.AutoFlush = true;
                        Log(log, w);
                    }
                    Console.Error.WriteLine("Failed to calculate plan dose. Please do it manually.");
                }
            }
//            if (isHalcyon && isArcBeamPlan)  // this plan has arc beams on a Halcyon or Ethos
            if (1 == 0)
            {
                Console.WriteLine("Halcyon arc plan.");
                var newPlan = extPlanSetup.Course.CopyPlanSetup(extPlanSetup);
                newPlan.Id = $"rot_{extPlanSetup.Id}";
                foreach (var bm in newPlan.Beams.Where(bm_ => !(bm_.IsSetupField || bm_.IsImagingTreatmentField)))
                {
                    var bPars = bm.GetEditableParameters();
                    foreach (var cpp in bPars.ControlPoints)
                    {
                        cpp.LeafPositions = Flipped(cpp.LeafPositions);
                    }
                    bm.ApplyParameters(bPars);
                }
            }
            if (isHalcyon && isStaticBeamPlan)  // this plan has static beams on a Halcyon or Ethos
            {
                Console.WriteLine("Halcyon static plan.");
                var newPlan = extPlanSetup.Course.CopyPlanSetup(extPlanSetup);
                newPlan.Id = $"inv_{extPlanSetup.Id}";
                foreach (var bm in newPlan.Beams.Where(bm_ => !(bm_.IsSetupField || bm_.IsImagingTreatmentField)))
                {
                    var bPars = bm.GetEditableParameters();
                    Console.WriteLine($"# of control points = {bPars.ControlPoints.Count()}");
                    int count = 0;
                    foreach (var cpp in bPars.ControlPoints)
                    {
                        count++;
                        float[,] leaves = cpp.LeafPositions;
                        if (count <= 2)
                        {
                            using (StreamWriter w = File.AppendText(logName))
                            {
                                string log = "";
                                for (int i = 0; i < leaves.Rank; i++)
                                {
                                    log += $"Array length at dimension {i}: {leaves.GetLength(i)}";
                                }
                                Log(log, w);
                                for (int i = 0; i < leaves.GetLength(0); i++)
                                {
                                    log = "";
                                    log += $"[{i}, x]: ";
                                    for (int j = 0; j < leaves.GetLength(1); j++)
                                    {
                                        log += $"{leaves[i, j]}, ";
                                    }
                                    Log(log, w);
                                }
                            }
                        }
                        cpp.LeafPositions = Rotated(cpp.LeafPositions);
                        leaves = cpp.LeafPositions;
                        if (count <= 2)
                        {
                            using (StreamWriter w = File.AppendText(logName))
                            {
                                Log("After:...", w);
                                string log = "";
                                for (int i = 0; i < leaves.Rank; i++)
                                {
                                    log += $"Array length at dimension {i}: {leaves.GetLength(i)}";
                                }
                                Log(log, w);
                                for (int i = 0; i < leaves.GetLength(0); i++)
                                {
                                log = "";
                                    log += $"[{i}, x]: ";
                                    for (int j = 0; j < leaves.GetLength(1); j++)
                                    {
                                        log += $"{leaves[i, j]}, ";
                                    }
                                Log(log, w);
                                }
                            }
                        }
                    }
                    bm.AddFlatteningSequence();
                    try
                    {
                        bm.ApplyParameters(bPars);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Exception: \n" + ex.ToString());
                    }
                }
            }
//            if (isSX2MLC && isStaticBeamPlan)  // this plan has static beams on a Halcyon or Ethos
            if (1 == 0)
            {
                ExternalBeamMachineParameters[] machineParameters = new ExternalBeamMachineParameters[numTxBeams];
                string[] machineID = new string[numTxBeams];
                string[] beamId = new string[numTxBeams];
                string[] newBeamId = new string[numTxBeams];
                string[] beamName = new string[numTxBeams];
                int[] doseRate = new int[numTxBeams];
                string[] energyModeId = new string[numTxBeams];
                string[] technique = new string[numTxBeams];
                string[] primaryFluenceModeId = new string[numTxBeams];
                MLCPlanType[] mlcType = new MLCPlanType[numTxBeams];
                double[] collimatorAngle = new double[numTxBeams];
                double[] gantryAngle = new double[numTxBeams];
                int[] numControlPoints = new int[numTxBeams];
                double[] patientSupportAngle = new double[numTxBeams];
                VVector[] isocenterPosition = new VVector[numTxBeams];
                VRect<double>[] jawPositions = new VRect<double>[numTxBeams];
                double[] weightFactor = new double[numTxBeams];
                float[][,] leafPositions = new float[numTxBeams][,];
                var cpsList = new ArrayList();
                int indexTxBeams = 0;
                foreach (Beam beam in extPlanSetup.Beams)
                {
                    if (beam.IsSetupField || beam.IsImagingTreatmentField)
                    {
                        continue;
                    }
                    using (StreamWriter w = File.AppendText(logName))  // record some log data
                    {
                        w.AutoFlush = true;
                        string log = $"Analyzing beam: {beam.Id}: \"{beam.Name}\". Technique: {beam.Technique.Id}. ";
                        log += $"It has {beam.ControlPoints.Count} control points.";
                        Log(log, w);
                        foreach (var cps in beam.ControlPoints)
                        {
                            log = $"Control Point {cps.Index}. Gantry = {cps.GantryAngle}; Collimator = {cps.CollimatorAngle}; ";
                            log += $"X1 Jaw = {cps.JawPositions.X1} mm; ";
                            log += $"X2 Jaw = {cps.JawPositions.X2} mm; ";
                            log += $"Y1 Jaw = {cps.JawPositions.Y1} mm; ";
                            log += $"Y2 Jaw = {cps.JawPositions.Y2} mm; ";
                            log += $"Meterweight = {cps.MetersetWeight}\n";
                            float[,] leaves = cps.LeafPositions;
                            log += $"{leaves.Length} leaf positions are found. ";
                            log += $"Leaf position array rank: {leaves.Rank}. ";
                            for (int i = 0; i < leaves.Rank; i++)
                            {
                                log += $"Array length at dimension {i}: {leaves.GetLength(i)}";
                            }
                            Log(log, w);
                            log = "";
                            for (int i = 0; i < leaves.GetLength(0); i++)
                            {
                                log += $"[{i}, x]: ";
                                for (int j = 0; j < leaves.GetLength(1); j++)
                                {
                                    log += $"{leaves[i, j]}, ";
                                }
                            }
                            Log(log, w);
                        }
                    }
                    ExternalBeamTreatmentUnit machine = beam.TreatmentUnit;
                    machineID[indexTxBeams] = machine.Id;
                    beamId[indexTxBeams] = beam.Id;
                    beamName[indexTxBeams] = beam.Name;
                    doseRate[indexTxBeams] = beam.DoseRate;
                    energyModeId[indexTxBeams] = beam.EnergyModeDisplayName;
                    technique[indexTxBeams] = beam.Technique.ToString();
                    mlcType[indexTxBeams] = beam.MLCPlanType;
                    weightFactor[indexTxBeams] = beam.WeightFactor;
                    if (energyModeId[indexTxBeams].Contains("-") && energyModeId[indexTxBeams].Split('-')[1] == "FFF")
                    {
                        primaryFluenceModeId[indexTxBeams] = energyModeId[indexTxBeams].Split('-')[1];
                        energyModeId[indexTxBeams] = energyModeId[indexTxBeams].Split('-')[0];
                    }
                    else
                    {
                        primaryFluenceModeId[indexTxBeams] = null;
                    }
                    machineParameters[indexTxBeams] = new ExternalBeamMachineParameters(machineID[indexTxBeams], energyModeId[indexTxBeams],
                        doseRate[indexTxBeams], technique[indexTxBeams], primaryFluenceModeId[indexTxBeams]);
                    collimatorAngle[indexTxBeams] = beam.ControlPoints[0].CollimatorAngle;
                    patientSupportAngle[indexTxBeams] = beam.ControlPoints[0].PatientSupportAngle;
                    numControlPoints[indexTxBeams] = beam.ControlPoints.Count;
                    isocenterPosition[indexTxBeams] = beam.IsocenterPosition;
                    jawPositions[indexTxBeams] = beam.ControlPoints[0].JawPositions;
                    leafPositions[indexTxBeams] = beam.ControlPoints[0].LeafPositions;
                    gantryAngle[indexTxBeams] = beam.ControlPoints[0].GantryAngle;
                    // Next we rotate the beam.
                    // First rotate the gantry for the beam.
                    if (gantryAngle[indexTxBeams] != 0)
                    {
                        gantryAngle[indexTxBeams] = 360 - gantryAngle[indexTxBeams];
                    }
                    // On a Varian machine with the SX2 MLC (Halcyon, Ethos), the collimator angle can only be in [270, 359.9] and [0, 90].
                    // So, the collimator cannot be rotated by 180 degrees.
                    // We always need to rotate the MLC aperture.
                    // The jaws are fixed with X1 = -14, X2 = 14, Y1 = -14, Y2 = 14. We do not need to change the jaws.

                    // In a Halcyon/Ethos plan, a static beam always has 300 control points (called fixed sequence) and
                    // the MLC plan type is always DoseDynamic.
                    // Like stated in the previous section, probably this "if" section will always run.
                    if (mlcType[indexTxBeams] == MLCPlanType.DoseDynamic)
                    {
                        List<CPModel> cpList = new List<CPModel>();
                        int indx = 0;
                        foreach (var cps in beam.ControlPoints)
                        {
                            float[,] currentLeaves = cps.LeafPositions;
                            float[,] newLeaves = new float[currentLeaves.GetLength(0), currentLeaves.GetLength(1)];
                            if (indx == 0)
                            {
                                Console.Write("i = 0: ");
                                for (int i = 0; i < 28; i++)
                                {
                                    Console.Write($"{currentLeaves[0, i]} ");
                                }
                                Console.WriteLine("");
                                Console.Write("i = 1: ");
                                for (int i = 0; i < 28; i++)
                                {
                                    Console.Write($"{currentLeaves[1, i]} ");
                                }
                                Console.WriteLine("");
                            }
                            indx++;
                            for (int i = 0; i < 28; i++)
                            {
                                newLeaves[0, i] = currentLeaves[0, i];
                                newLeaves[1, i] = currentLeaves[1, i];
                                //                                if (currentLeaves[1, 28 - i - 1] != 140 && currentLeaves[1, 28 - i - 1] != -140)
                                //                                {
                                //                                    newLeaves[0, i] = -currentLeaves[1, 28 - i - 1];
                                //                                }
                                //                                else
                                //{
                                //newLeaves[0, i] = currentLeaves[1, 28 - i - 1];
                                //}
                                //if (currentLeaves[0, i] != 140 && currentLeaves[0, i] != -140)
                                //{
                                //newLeaves[1, 28 - i - 1] = -currentLeaves[0, i];
                                //}
                                //                                else
                                //{
                                //newLeaves[1, 28 - i - 1] = currentLeaves[0, i];
                                //}
                            }
                            for (int i = 28; i < 57; i++)
                            {
                                newLeaves[0, i] = currentLeaves[0, i];
                                newLeaves[1, i] = currentLeaves[1, i];
                            }
                            cpList.Add(new CPModel(57)
                            {
                                GantryAngle = cps.GantryAngle,
                                JawPositions = new VRect<double>(-cps.JawPositions.X2, -cps.JawPositions.Y2, -cps.JawPositions.X1, -cps.JawPositions.Y1),
                                MetersetWeight = cps.MetersetWeight,
                                CollimatorAngle = cps.CollimatorAngle,
                                MLCPositions = newLeaves
                            });
                        }
                        cpsList.Add(cpList);
                    }
                    indexTxBeams++;
                }
                // create new beams with rotated geometry
                for (int idxTxBeam = 0; idxTxBeam < numTxBeams; idxTxBeam++)
                {
                    using (StreamWriter w = File.AppendText(logName))
                    {
                        w.AutoFlush = true;
                        string log = $"Creating a new rotated beam for beam: \"{beamId[idxTxBeam]}\"";
                        Log(log, w);
                    }
                    if (mlcType[idxTxBeam] == MLCPlanType.Static)  // this beam has a static MLC.
                    {
                        Beam newBeam = extPlanSetup.AddMLCBeam(machineParameters[idxTxBeam], leafPositions[idxTxBeam], jawPositions[idxTxBeam],
                            collimatorAngle[idxTxBeam], gantryAngle[idxTxBeam], patientSupportAngle[idxTxBeam], isocenterPosition[idxTxBeam]);
                        BeamParameters bParam = newBeam.GetEditableParameters();
                        bParam.WeightFactor = weightFactor[idxTxBeam];
                        bParam.SetAllLeafPositions(leafPositions[idxTxBeam]);
                        newBeam.ApplyParameters(bParam);
                        newBeam.Name = beamName[idxTxBeam];
                        newBeamId[idxTxBeam] = newBeam.Id;
                    }
                    else if (mlcType[idxTxBeam] == MLCPlanType.DoseDynamic)  // this field uses dynamic MLC positions.
                    {
                        isDynamicBeamPlan = true;
                        // ESAPI manual shows that the number of meterset weight items define the number of created control points.
                        // A Halcyon static beam has 300 control points. Each control point increases the weight by 1/300.
                        List<CPModel> cpList = (List<CPModel>)cpsList[idxTxBeam];
                        Beam newBeam = extPlanSetup.AddFixedSequenceBeam(machineParameters[idxTxBeam], collimatorAngle[idxTxBeam],
                            gantryAngle[idxTxBeam], isocenterPosition[idxTxBeam]);
                        Console.WriteLine($"Beam index: {idxTxBeam}");
                        Console.WriteLine($"New Beam ID: {newBeam.Id}");
                        Console.WriteLine($"MLC plan type: {newBeam.MLCPlanType}");
                        Console.WriteLine($"Collimator Angle: {collimatorAngle[idxTxBeam]}");
                        Console.WriteLine($"Gantry Angle: {gantryAngle[idxTxBeam]}");
                        BeamParameters bParam = newBeam.GetEditableParameters();
                        Console.WriteLine($"# Control points: {bParam.ControlPoints.Count()}");
                        for (int indCP = 0; indCP < cpList.Count; indCP++)
                        {
//                            bParam.ControlPoints.ElementAt(indCP).JawPositions = cpList[indCP].JawPositions;
                            bParam.ControlPoints.ElementAt(indCP).LeafPositions = cpList[indCP].MLCPositions;
                            //                            for (int i = 0; i < bParam.ControlPoints.ElementAt(indCP).LeafPositions.GetLength(0); i++)
                            //                            {
                            //Console.Write($"Index {indCP}\ti={i} ");
                            //for (int j = 0; j < 57; j++)
                            //{
                            //Console.Write($"{bParam.ControlPoints.ElementAt(indCP).LeafPositions[i, j]}, ");
                            //}
                            //Console.WriteLine("");
                            //}
                        }
                        bParam.WeightFactor = weightFactor[idxTxBeam];
//                        Console.WriteLine($"{newBeam.AddFlatteningSequence()}");
                        Console.WriteLine($"{bParam.ControlPoints.Count()}");
                        try
                        {
                            newBeam.ApplyParameters(bParam);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.ToString());
                        }
                        newBeam.Name = beamName[idxTxBeam];
                        newBeamId[idxTxBeam] = newBeam.Id;
                    }
                }
                // Remove the old treatment beams.
                foreach (string eachTxBeamId in beamId)
                {
                    try
                    {
                        extPlanSetup.RemoveBeam(extPlanSetup.Beams.Where(b => b.Id == eachTxBeamId).Single());
                    }
                    catch (Exception e)
                    {
                        string log = e.ToString();
                        using (StreamWriter w = File.AppendText(logName))
                        {
                            w.AutoFlush = true;
                            Log(log, w);
                        }
                        Console.Error.WriteLine($"Failed to remove one existing beam: \"{eachTxBeamId}\". Please remove it manually.");
                    }
                }
                // Calculate plan dose with the new beams
                try
                {
                    if (isDynamicBeamPlan == false)
                    {
                        extPlanSetup.CalculateDose();
                    }
                    else
                    {
                        // first create a <beamID, beam meterset> list as preset values.
                        List<KeyValuePair<string, MetersetValue>> presetValues = new List<KeyValuePair<string, MetersetValue>>();
                        for (int idxTxBeam = 0; idxTxBeam < numTxBeams; idxTxBeam++)
                        {
                            presetValues.Add(new KeyValuePair<string, MetersetValue>(newBeamId[idxTxBeam],
                                extPlanSetup.Beams.FirstOrDefault(b => b.Id == beamId[idxTxBeam]).Meterset));
                        }
                        extPlanSetup.CalculateDoseWithPresetValues(presetValues);
                    }
                }
                catch (Exception e)
                {
                    string log = e.ToString();
                    using (StreamWriter w = File.AppendText(logName))
                    {
                        w.AutoFlush = true;
                        Log(log, w);
                    }
                    Console.Error.WriteLine("Failed to calculate plan dose. Please do it manually.");
                }
            }
            if (isSX2MLC == false && isArcBeamPlan)
            {
                ExternalBeamMachineParameters[] machineParameters = new ExternalBeamMachineParameters[numTxBeams];
                string[] machineID = new string[numTxBeams];
                string[] beamId = new string[numTxBeams];
                string[] newBeamId = new string[numTxBeams];
                string[] beamName = new string[numTxBeams];
                int[] doseRate = new int[numTxBeams];
                string[] energyModeId = new string[numTxBeams];
                string[] primaryFluenceModeId = new string[numTxBeams];
                string[] technique = new string[numTxBeams];
                MLCPlanType[] mlcType = new MLCPlanType[numTxBeams];
                double[] collimatorAngle = new double[numTxBeams];
                double[] gantryStartAngle = new double[numTxBeams];
                double[] gantryStopAngle = new double[numTxBeams];
                GantryDirection[] gantryDirection = new GantryDirection[numTxBeams];
                int[] numControlPoints = new int[numTxBeams];
                double[] patientSupportAngle = new double[numTxBeams];
                VVector[] isocenterPosition = new VVector[numTxBeams];
                VRect<double>[] jawPositions = new VRect<double>[numTxBeams];
                double[] weightFactor = new double[numTxBeams];
                float[][,] leafPositions = new float[numTxBeams][,];
                var cpsList = new ArrayList();
                int indexTxBeams = 0;
                foreach (Beam beam in extPlanSetup.Beams)
                {
                    if (beam.IsSetupField)
                    {
                        continue;
                    }
                    using (StreamWriter w = File.AppendText(logName))  // record some log data
                    {
                        w.AutoFlush = true;
                        string log = $"Analyzing beam: {beam.Id}: \"{beam.Name}\". Technique: {beam.Technique.Id}. ";
                        log += $"It has {beam.ControlPoints.Count} control points.";
                        Log(log, w);
                        foreach (var cps in beam.ControlPoints)
                        {
                            log = $"Control Point {cps.Index}. Gantry = {cps.GantryAngle}; Collimator = {cps.CollimatorAngle}; ";
                            log += $"X1 Jaw = {cps.JawPositions.X1} mm; ";
                            log += $"X2 Jaw = {cps.JawPositions.X2} mm; ";
                            log += $"Y1 Jaw = {cps.JawPositions.Y1} mm; ";
                            log += $"Y2 Jaw = {cps.JawPositions.Y2} mm; ";
                            log += $"Meterweight = {cps.MetersetWeight}\n";
                            float[,] leaves = cps.LeafPositions;
                            log += $"{leaves.Length} leaf positions are found. ";
                            log += $"Leaf position array rank: {leaves.Rank}. ";
                            for (int i = 0; i < leaves.Rank; i++)
                            {
                                log += $"Array length at dimension {i}: {leaves.GetLength(i)}";
                            }
                            Log(log, w);
                            log = "";
                            for (int i = 0; i < leaves.GetLength(0); i++)
                            {
                                log += $"[{i}, x]: ";
                                for (int j = 0; j < leaves.GetLength(1); j++)
                                {
                                    log += $"{leaves[i, j]}, ";
                                }
                            }
                            Log(log, w);
                        }
                    }
                    ExternalBeamTreatmentUnit machine = beam.TreatmentUnit;
                    machineID[indexTxBeams] = machine.Id;
                    beamId[indexTxBeams] = beam.Id;
                    beamName[indexTxBeams] = beam.Name;
                    doseRate[indexTxBeams] = beam.DoseRate;
                    energyModeId[indexTxBeams] = beam.EnergyModeDisplayName;
                    technique[indexTxBeams] = beam.Technique.ToString();
                    if (energyModeId[indexTxBeams].Contains("-") && energyModeId[indexTxBeams].Split('-')[1] == "FFF")
                    {
                        primaryFluenceModeId[indexTxBeams] = energyModeId[indexTxBeams].Split('-')[1];
                        energyModeId[indexTxBeams] = energyModeId[indexTxBeams].Split('-')[0];
                    }
                    else
                    {
                        primaryFluenceModeId[indexTxBeams] = null;
                    }
                    mlcType[indexTxBeams] = beam.MLCPlanType;
                    weightFactor[indexTxBeams] = beam.WeightFactor;
                    machineParameters[indexTxBeams] = new ExternalBeamMachineParameters(machineID[indexTxBeams], energyModeId[indexTxBeams],
                        doseRate[indexTxBeams], technique[indexTxBeams], primaryFluenceModeId[indexTxBeams]);
                    collimatorAngle[indexTxBeams] = beam.ControlPoints[0].CollimatorAngle;
                    patientSupportAngle[indexTxBeams] = beam.ControlPoints[0].PatientSupportAngle;
                    numControlPoints[indexTxBeams] = beam.ControlPoints.Count;
                    isocenterPosition[indexTxBeams] = beam.IsocenterPosition;
                    jawPositions[indexTxBeams] = beam.ControlPoints[0].JawPositions;
                    leafPositions[indexTxBeams] = beam.ControlPoints[0].LeafPositions;
                    gantryStartAngle[indexTxBeams] = beam.ControlPoints.First().GantryAngle;
                    gantryStopAngle[indexTxBeams] = beam.ControlPoints.Last().GantryAngle;
                    gantryDirection[indexTxBeams] = beam.GantryDirection;
                    // Next we rotate the beam.
                    // First rotate the gantry for the beam.
                    if (gantryStartAngle[indexTxBeams] != 0)
                    {
                        gantryStartAngle[indexTxBeams] = 360 - gantryStartAngle[indexTxBeams];
                    }
                    if (gantryStopAngle[indexTxBeams] != 0)
                    {
                        gantryStopAngle[indexTxBeams] = 360 - gantryStopAngle[indexTxBeams];
                    }
                    if (gantryDirection[indexTxBeams] == GantryDirection.Clockwise)
                    {
                        gantryDirection[indexTxBeams] = GantryDirection.CounterClockwise;
                    }
                    else if (gantryDirection[indexTxBeams] == GantryDirection.CounterClockwise)
                    {
                        gantryDirection[indexTxBeams] = GantryDirection.Clockwise;
                    }
                    // On a Varian TrueBeam, the collimator angle cannot be in (175, 185).
                    // if the collimator angle is within (355, 5), we cannot rotate the collimator to the opposite side.
                    if (collimatorAngle[indexTxBeams] >= 5 && collimatorAngle[indexTxBeams] <= 355)
                    {
                        collimatorAngle[indexTxBeams] += 180;
                        if (collimatorAngle[indexTxBeams] >= 360)
                        {
                            collimatorAngle[indexTxBeams] = collimatorAngle[indexTxBeams] - 360;
                        }
                        List<CPModel> cpList = new List<CPModel>();
                        foreach (var cps in beam.ControlPoints)
                        {
                            int numLeafPairs = cps.LeafPositions.GetLength(1);
                            cpList.Add(new CPModel(numLeafPairs)
                            {
                                GantryAngle = cps.GantryAngle,
                                JawPositions = cps.JawPositions,
                                MetersetWeight = cps.MetersetWeight,
                                CollimatorAngle = cps.CollimatorAngle,
                                MLCPositions = cps.LeafPositions
                            });
                        }
                        cpsList.Add(cpList);
                    }
                    else  // we rotate the jaw and MLC aperture here instead of the collimator.
                    {
                        VRect<double> currentJaws = jawPositions[indexTxBeams];
                        VRect<double> newJaws = new VRect<double>(-currentJaws.X2, -currentJaws.Y2, -currentJaws.X1, -currentJaws.Y1);
                        jawPositions[indexTxBeams] = newJaws;
                        List<CPModel> cpList = new List<CPModel>();
                        foreach (var cps in beam.ControlPoints)
                        {
                            float[,] currentLeaves = cps.LeafPositions;
                            float[,] newLeaves = new float[currentLeaves.GetLength(0), currentLeaves.GetLength(1)];
                            for (int i = 0; i < currentLeaves.GetLength(1); i++)
                            {
                                newLeaves[0, i] = -currentLeaves[1, currentLeaves.GetLength(1) - i - 1];
                                newLeaves[1, currentLeaves.GetLength(1) - i - 1] = -currentLeaves[0, i];
                            }
                            cpList.Add(new CPModel
                            {
                                GantryAngle = cps.GantryAngle,
                                JawPositions = new VRect<double>(-cps.JawPositions.X2, -cps.JawPositions.Y2, -cps.JawPositions.X1, -cps.JawPositions.Y1),
                                MetersetWeight = cps.MetersetWeight,
                                CollimatorAngle = cps.CollimatorAngle,
                                MLCPositions = newLeaves
                            });
                        }
                        cpsList.Add(cpList);
                    }
                    indexTxBeams++;
                }
                // create new beams with rotated geometry
                for (int idxTxBeam = 0; idxTxBeam < numTxBeams; idxTxBeam++)
                {
                    using (StreamWriter w = File.AppendText(logName))
                    {
                        w.AutoFlush = true;
                        string log = $"Creating a new rotated beam for beam: \"{beamId[idxTxBeam]}\"";
                        Log(log, w);
                    }
                    // ESAPI manual shows that the number of meterset weight items define the number of created control points.
                    List<CPModel> cpList = (List<CPModel>)cpsList[idxTxBeam];
                    var metersetWeights = from cp in cpList select cp.MetersetWeight;
                    Beam newBeam = extPlanSetup.AddVMATBeam(machineParameters[idxTxBeam], metersetWeights, collimatorAngle[idxTxBeam],
                        gantryStartAngle[idxTxBeam], gantryStopAngle[idxTxBeam], gantryDirection[idxTxBeam],
                        patientSupportAngle[idxTxBeam], isocenterPosition[idxTxBeam]);
                    BeamParameters bParam = newBeam.GetEditableParameters();
                    for (int indCP = 0; indCP < cpList.Count; indCP++)
                    {
                        bParam.ControlPoints.ElementAt(indCP).JawPositions = cpList[indCP].JawPositions;
                        bParam.ControlPoints.ElementAt(indCP).LeafPositions = cpList[indCP].MLCPositions;
                    }
                    bParam.WeightFactor = weightFactor[idxTxBeam];
                    newBeam.ApplyParameters(bParam);
                    newBeam.Name = beamName[idxTxBeam];
                    newBeamId[idxTxBeam] = newBeam.Id;
                }
                // Remove the old treatment beams.
                foreach (string eachTxBeamId in beamId)
                {
                    try
                    {
                        extPlanSetup.RemoveBeam(extPlanSetup.Beams.Where(b => b.Id == eachTxBeamId).Single());
                    }
                    catch (Exception e)
                    {
                        string log = e.ToString();
                        using (StreamWriter w = File.AppendText(logName))
                        {
                            w.AutoFlush = true;
                            Log(log, w);
                        }
                        Console.Error.WriteLine($"Failed to remove one existing beam: \"{eachTxBeamId}\". Please remove it manually.");
                    }
                }
                // Calculate plan dose with the new beams
                try
                {
                    // first create a <beamID, beam meterset> list as preset values.
                    List<KeyValuePair<string, MetersetValue>> presetValues = new List<KeyValuePair<string, MetersetValue>>();
                    // create new beams with rotated geometry
                    for (int idxTxBeam = 0; idxTxBeam < numTxBeams; idxTxBeam++)
                    {
                        presetValues.Add(new KeyValuePair<string, MetersetValue>(newBeamId[idxTxBeam],
                            extPlanSetup.Beams.FirstOrDefault(b => b.Id == beamId[idxTxBeam]).Meterset));
                    }
                    extPlanSetup.CalculateDoseWithPresetValues(presetValues);
                }
                catch (Exception e)
                {
                    string log = e.ToString();
                    using (StreamWriter w = File.AppendText(logName))
                    {
                        w.AutoFlush = true;
                        Log(log, w);
                    }
                    Console.Error.WriteLine("Failed to calculate plan dose. Please do it manually.");
                }
            }
            if (isSX2MLC == false && isStaticBeamPlan)
            {
                ExternalBeamMachineParameters[] machineParameters = new ExternalBeamMachineParameters[numTxBeams];
                string[] machineID = new string[numTxBeams];
                string[] beamId = new string[numTxBeams];
                string[] newBeamId = new string[numTxBeams];
                string[] beamName = new string[numTxBeams];
                int[] doseRate = new int[numTxBeams];
                string[] energyModeId = new string[numTxBeams];
                string[] technique = new string[numTxBeams];
                string[] primaryFluenceModeId = new string[numTxBeams];
                MLCPlanType[] mlcType = new MLCPlanType[numTxBeams];
                double[] collimatorAngle = new double[numTxBeams];
                double[] gantryAngle = new double[numTxBeams];
                int[] numControlPoints = new int[numTxBeams];
                double[] patientSupportAngle = new double[numTxBeams];
                VVector[] isocenterPosition = new VVector[numTxBeams];
                VRect<double>[] jawPositions = new VRect<double>[numTxBeams];
                double[] weightFactor = new double[numTxBeams];
                float[][,] leafPositions = new float[numTxBeams][,];
                var cpsList = new ArrayList();
                int indexTxBeams = 0;
                foreach (Beam beam in extPlanSetup.Beams)
                {
                    if (beam.IsSetupField)
                    {
                        continue;
                    }
                    using (StreamWriter w = File.AppendText(logName))  // record some log data
                    {
                        w.AutoFlush = true;
                        string log = $"Analyzing beam: {beam.Id}: \"{beam.Name}\". Technique: {beam.Technique.Id}. ";
                        log += $"It has {beam.ControlPoints.Count} control points.";
                        Log(log, w);
                        foreach (var cps in beam.ControlPoints)
                        {
                            log = $"Control Point {cps.Index}. Gantry = {cps.GantryAngle}; Collimator = {cps.CollimatorAngle}; ";
                            log += $"X1 Jaw = {cps.JawPositions.X1} mm; ";
                            log += $"X2 Jaw = {cps.JawPositions.X2} mm; ";
                            log += $"Y1 Jaw = {cps.JawPositions.Y1} mm; ";
                            log += $"Y2 Jaw = {cps.JawPositions.Y2} mm; ";
                            log += $"Meterweight = {cps.MetersetWeight}\n";
                            float[,] leaves = cps.LeafPositions;
                            log += $"{leaves.Length} leaf positions are found. ";
                            log += $"Leaf position array rank: {leaves.Rank}. ";
                            for (int i = 0; i < leaves.Rank; i++)
                            {
                                log += $"Array length at dimension {i}: {leaves.GetLength(i)}";
                            }
                            Log(log, w);
                            log = "";
                            for (int i = 0; i < leaves.GetLength(0); i++)
                            {
                                log += $"[{i}, x]: ";
                                for (int j = 0; j < leaves.GetLength(1); j++)
                                {
                                    log += $"{leaves[i, j]}, ";
                                }
                            }
                            Log(log, w);
                        }
                    }
                    ExternalBeamTreatmentUnit machine = beam.TreatmentUnit;
                    machineID[indexTxBeams] = machine.Id;
                    beamId[indexTxBeams] = beam.Id;
                    beamName[indexTxBeams] = beam.Name;
                    doseRate[indexTxBeams] = beam.DoseRate;
                    energyModeId[indexTxBeams] = beam.EnergyModeDisplayName;
                    technique[indexTxBeams] = beam.Technique.ToString();
                    mlcType[indexTxBeams] = beam.MLCPlanType;
                    weightFactor[indexTxBeams] = beam.WeightFactor;
                    if (energyModeId[indexTxBeams].Contains("-") && energyModeId[indexTxBeams].Split('-')[1] == "FFF")
                    {
                        primaryFluenceModeId[indexTxBeams] = energyModeId[indexTxBeams].Split('-')[1];
                        energyModeId[indexTxBeams] = energyModeId[indexTxBeams].Split('-')[0];
                    }
                    else
                    {
                        primaryFluenceModeId[indexTxBeams] = null;
                    }
                    machineParameters[indexTxBeams] = new ExternalBeamMachineParameters(machineID[indexTxBeams], energyModeId[indexTxBeams],
                        doseRate[indexTxBeams], technique[indexTxBeams], primaryFluenceModeId[indexTxBeams]);
                    collimatorAngle[indexTxBeams] = beam.ControlPoints[0].CollimatorAngle;
                    patientSupportAngle[indexTxBeams] = beam.ControlPoints[0].PatientSupportAngle;
                    numControlPoints[indexTxBeams] = beam.ControlPoints.Count;
                    isocenterPosition[indexTxBeams] = beam.IsocenterPosition;
                    jawPositions[indexTxBeams] = beam.ControlPoints[0].JawPositions;
                    leafPositions[indexTxBeams] = beam.ControlPoints[0].LeafPositions;
                    gantryAngle[indexTxBeams] = beam.ControlPoints[0].GantryAngle;
                    // Next we rotate the beam.
                    // First rotate the gantry for the beam.
                    if (gantryAngle[indexTxBeams] != 0)
                    {
                        gantryAngle[indexTxBeams] = 360 - gantryAngle[indexTxBeams];
                    }
                    // On a Varian TrueBeam, the collimator angle cannot be in (175, 185).
                    // if the collimator angle is within (355, 5), we cannot rotate the collimator to the opposite side.
                    if (collimatorAngle[indexTxBeams] >= 5 && collimatorAngle[indexTxBeams] <= 355)
                    {
                        collimatorAngle[indexTxBeams] += 180;
                        if (collimatorAngle[indexTxBeams] >= 360)
                        {
                            collimatorAngle[indexTxBeams] = collimatorAngle[indexTxBeams] - 360;
                        }
                        List<CPModel> cpList = new List<CPModel>();
                        foreach (var cps in beam.ControlPoints)
                        {
                            cpList.Add(new CPModel
                            {
                                GantryAngle = cps.GantryAngle,
                                JawPositions = cps.JawPositions,
                                MetersetWeight = cps.MetersetWeight,
                                CollimatorAngle = cps.CollimatorAngle,
                                MLCPositions = cps.LeafPositions
                            });
                        }
                        cpsList.Add(cpList);
                    }
                    else  // we rotate the jaw and MLC aperture here instead of the collimator.
                    {
                        VRect<double> currentJaws = jawPositions[indexTxBeams];
                        VRect<double> newJaws = new VRect<double>(-currentJaws.X2, -currentJaws.Y2, -currentJaws.X1, -currentJaws.Y1);
                        jawPositions[indexTxBeams] = newJaws;
                        if (mlcType[indexTxBeams] == MLCPlanType.Static)
                        {
                            float[,] currentLeaves = leafPositions[indexTxBeams];
                            float[,] newLeaves = new float[currentLeaves.GetLength(0), currentLeaves.GetLength(1)];
                            for (int i = 0; i < currentLeaves.GetLength(1); i++)
                            {
                                newLeaves[0, i] = -currentLeaves[1, currentLeaves.GetLength(1) - i - 1];
                                newLeaves[1, currentLeaves.GetLength(1) - i - 1] = -currentLeaves[0, i];
                            }
                            leafPositions[indexTxBeams] = newLeaves;
                        }
                        else if (mlcType[indexTxBeams] == MLCPlanType.DoseDynamic)
                        {
                            List<CPModel> cpList = new List<CPModel>();
                            foreach (var cps in beam.ControlPoints)
                            {
                                float[,] currentLeaves = cps.LeafPositions;
                                float[,] newLeaves = new float[currentLeaves.GetLength(0), currentLeaves.GetLength(1)];
                                for (int i = 0; i < currentLeaves.GetLength(1); i++)
                                {
                                    newLeaves[0, i] = -currentLeaves[1, currentLeaves.GetLength(1) - i - 1];
                                    newLeaves[1, currentLeaves.GetLength(1) - i - 1] = -currentLeaves[0, i];
                                }
                                cpList.Add(new CPModel
                                {
                                    GantryAngle = cps.GantryAngle,
                                    JawPositions = new VRect<double>(-cps.JawPositions.X2, -cps.JawPositions.Y2, -cps.JawPositions.X1, -cps.JawPositions.Y1),
                                    MetersetWeight = cps.MetersetWeight,
                                    CollimatorAngle = cps.CollimatorAngle,
                                    MLCPositions = newLeaves
                                });
                            }
                            cpsList.Add(cpList);

                        }
                    }
                    indexTxBeams++;
                }
                // create new beams with rotated geometry
                for(int idxTxBeam = 0; idxTxBeam < numTxBeams; idxTxBeam++)
                {
                    using (StreamWriter w = File.AppendText(logName))
                    {
                        w.AutoFlush = true;
                        string log = $"Creating a new rotated beam for beam: \"{beamId[idxTxBeam]}\"";
                        Log(log, w);
                    }
                    if (mlcType[idxTxBeam] == MLCPlanType.NotDefined)  // this beam does not have an MLC.
                    {
                        Beam newBeam = extPlanSetup.AddStaticBeam(machineParameters[idxTxBeam], jawPositions[idxTxBeam], 
                            collimatorAngle[idxTxBeam], gantryAngle[idxTxBeam], patientSupportAngle[idxTxBeam], isocenterPosition[idxTxBeam]);
                        BeamParameters bParam = newBeam.GetEditableParameters();
                        bParam.WeightFactor = weightFactor[idxTxBeam];
                        bParam.SetAllLeafPositions(leafPositions[idxTxBeam]);
                        newBeam.ApplyParameters(bParam);
                        newBeam.Name = beamName[idxTxBeam];
                        newBeamId[idxTxBeam] = newBeam.Id;
                    }
                    else if (mlcType[idxTxBeam] == MLCPlanType.Static)  // this beam has a static MLC.
                    {
                        Beam newBeam = extPlanSetup.AddMLCBeam(machineParameters[idxTxBeam], leafPositions[idxTxBeam], jawPositions[idxTxBeam],
                            collimatorAngle[idxTxBeam], gantryAngle[idxTxBeam], patientSupportAngle[idxTxBeam], isocenterPosition[idxTxBeam]);
                        BeamParameters bParam = newBeam.GetEditableParameters();
                        bParam.WeightFactor = weightFactor[idxTxBeam];
                        bParam.SetAllLeafPositions(leafPositions[idxTxBeam]);
                        newBeam.ApplyParameters(bParam);
                        newBeam.Name = beamName[idxTxBeam];
                        newBeamId[idxTxBeam] = newBeam.Id;
                    }
                    else if (mlcType[idxTxBeam] == MLCPlanType.DoseDynamic)  // this field uses dynamic MLC positions.
                    {
                        isDynamicBeamPlan = true;
                        // ESAPI manual shows that the number of meterset weight items define the number of created control points.
                        List<CPModel> cpList = (List<CPModel>) cpsList[idxTxBeam];
                        var metersetWeights = from cp in cpList select cp.MetersetWeight;
                        Beam newBeam = extPlanSetup.AddMultipleStaticSegmentBeam(machineParameters[idxTxBeam], metersetWeights, collimatorAngle[idxTxBeam],
                            gantryAngle[idxTxBeam], patientSupportAngle[idxTxBeam], isocenterPosition[idxTxBeam]);
                        BeamParameters bParam = newBeam.GetEditableParameters();
                        for (int indCP = 0; indCP < cpList.Count; indCP++)
                        {
                            bParam.ControlPoints.ElementAt(indCP).JawPositions = cpList[indCP].JawPositions;
                            bParam.ControlPoints.ElementAt(indCP).LeafPositions = cpList[indCP].MLCPositions;
                        }
                        bParam.WeightFactor = weightFactor[idxTxBeam];
                        newBeam.ApplyParameters(bParam);
                        newBeam.Name = beamName[idxTxBeam];
                        newBeamId[idxTxBeam] = newBeam.Id;
                    }
                }
                // Remove the old treatment beams.
                foreach (string eachTxBeamId in beamId)
                {
                    try
                    {
                        extPlanSetup.RemoveBeam(extPlanSetup.Beams.Where(b => b.Id == eachTxBeamId).Single());
                    }
                    catch (Exception e)
                    {
                        string log = e.ToString();
                        using (StreamWriter w = File.AppendText(logName))
                        {
                            w.AutoFlush = true;
                            Log(log, w);
                        }
                        Console.Error.WriteLine($"Failed to remove one existing beam: \"{eachTxBeamId}\". Please remove it manually.");
                    }
                }
                // Calculate plan dose with the new beams
                try
                {
                    if (isDynamicBeamPlan == false)
                    {
                        extPlanSetup.CalculateDose();
                    }
                    else
                    {
                        // first create a <beamID, beam meterset> list as preset values.
                        List<KeyValuePair<string, MetersetValue>> presetValues = new List<KeyValuePair<string, MetersetValue>>();
                        // create new beams with rotated geometry
                        for (int idxTxBeam = 0; idxTxBeam < numTxBeams; idxTxBeam++)
                        {
                            presetValues.Add(new KeyValuePair<string, MetersetValue>(newBeamId[idxTxBeam], 
                                extPlanSetup.Beams.FirstOrDefault(b => b.Id == beamId[idxTxBeam]).Meterset));
                        }
                            extPlanSetup.CalculateDoseWithPresetValues(presetValues);
                    }
                }
                catch(Exception e)
                {
                    string log = e.ToString();
                    using (StreamWriter w = File.AppendText(logName))
                    {
                        w.AutoFlush = true;
                        Log(log, w);
                    }
                    Console.Error.WriteLine("Failed to calculate plan dose. Please do it manually.");
                }
            }
            app.SaveModifications();
            app.ClosePatient();
        }
    }
    public class CPModel
    { 
        public double MetersetWeight { get; set; }
        public VRect<double> JawPositions { get; set; }
        public double GantryAngle { get; set; }
        public double CollimatorAngle { get; set; }
        public float[,] MLCPositions { get; set; }
        public CPModel(int numLeaves)
        {
            MLCPositions = new float[2, numLeaves];
        }
        public CPModel()
        {
            // If the number of leaves is not provided, we assume that the MLC has 60 leave pairs.
            MLCPositions = new float[2, 60];
        }
    }
}
