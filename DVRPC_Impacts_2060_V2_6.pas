program DVRPC_Impacts_2060_V2_6;


{$APPTYPE CONSOLE}

uses
  Classes, sysutils, fpstypes, fpSpreadsheet, laz_fpspreadsheet;

type TDynFloat = array of array of Real;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  datArray : TDynFloat;

function ReadSpreadsheetRange
         (MyWorksheet: TsWorksheet; startCell : string; endCell: string = '')
                       : TDynFloat;
// Function reads numbers FROM range of cells (startCell : endCell)
// and RETURNS them as a dynamical array of Floats.
// If endCell is not provided, it reads all cells from startCell to the last.

// Example usage:
// ReadSpreadsheetRange(MyWorksheet,'B2', 'F3')

const xdebug=false;
 var
   LastColumn, LastRow, row, col : integer;
   cell : PCell;
   myArray : TDynFloat;
 begin
   if MyWorksheet.FindCell(startCell) = nil then
      begin
        // ToDo: Maybe I should return an empty array before exit (???)
        WriteLn('Cell ' + startCell + ' does not exist in ' + MyWorksheet.Name);
        exit;
      end;

   cell := MyWorksheet.FindCell(startCell);

   // Set last column and rows depending on whether endCell provided
   if ( endCell = '' ) or ( MyWorksheet.FindCell(endCell) = nil ) then
      begin
        LastColumn := MyWorksheet.GetLastColIndex();
        LastRow := MyWorksheet.GetLastRowIndex();
      end
   else
     begin
       LastColumn := MyWorksheet.FindCell(endCell)^.Col;
       LastRow := MyWorksheet.FindCell(endCell)^.Row;
     end;

   SetLength(myArray,LastRow - cell^.Row + 1,LastColumn - cell^.Col + 1);
   for row := cell^.Row to LastRow do
       for col := cell^.Col to LastColumn do
           begin
             myArray[row - cell^.Row, col - cell^.Col] :=
                         MyWorksheet.ReadAsNumber(row,col);
           end;

  if xdebug then begin
    for row:=0 to LastRow-cell^.Row do begin
      for col:=0 to LastColumn-cell^.Col do write(myArray[row,col]:8:1);
      writeln;
    end;
    readln;
  end;

   Result := myArray;
 end;


function Dummy(a,b:integer): single;
begin
  if a=b then Dummy:=1.0 else Dummy:=0.0;
end;
function DummyRange(a,b,c:single): single;
begin
  if (a>=b) and (a<c) then DummyRange:=1.0 else DummyRange:=0.0;
end;

function Max(a,b:single): single;
begin
  if a>b then Max:=a else Max:=b;
end;

function Min(a,b:single): single;
begin
  if a<b then Min:=a else Min:=b;
end;

{Control Module}
{constants}
const
StartYear:single = 2010;
TimeStepLength = 0.5; {years}
NumberOfTimeSteps = 100;
NumberOfRegions = 1;
testWriteYear = 0;

{global variables}
var
outest:text;
TimeStep:integer = 0;
Year:single;

Region:integer = 1;
Scenario:integer = 1;

{Demographic Module}
{constants}
const
NumberOfDemographicDimensions = 6;

NumberOfAgeGroups = 17;
AgeGroupLabels:array[0..NumberOfAgeGroups] of string[29]=
('Total',
 'Age  0- 4',
 'Age  5- 9',
 'Age 10-15',
 'Age 16-19',
 'Age 20-24',
 'Age 25-29',
 'Age 30-34',
 'Age 35-39',
 'Age 40-44',
 'Age 45-49',
 'Age 50-54',
 'Age 55-59',
 'Age 60-64',
 'Age 65-69',
 'Age 70-74',
 'Age 75-79',
 'Age 80 up');

  BirthAgeGroup = 1; {new births go into youngest age group}
  AgeGroupDuration:array[1..NumberOfAgeGroups] of single = (5,5,6, 4,5,5, 5,5,5, 5,5,5, 5,5,5, 5,0);

NumberOfHhldTypes = 4;
HhldTypeLabels:array[1..NumberOfHhldTypes] of string[29]=
('Single/No Kids',
 'Couple/No Kids',
 'Single/With Kids',
 'Couple/With Kids');

  BirthHhldType:array[1..NumberOfHhldTypes] of integer=(2,4,2,4); {new births change 0 Ch to 1+ Ch}
  NumberOfAdults  :array[1..NumberOfHhldTypes] of single=(1, 2.2,   1, 2.2);
  NumberOfChildren:array[1..NumberOfHhldTypes] of single=(0,   0, 1.5, 1.5);

  NumberOfEthnicGrs = 12;
EthnicGrLabels:array[1..NumberOfEthnicGrs] of string[29]=
('Hispanic US born',
 'Hispanic >20 yrs',
 'Hispanic <20 yrs',
 'Black US born',
 'Black >20 yrs',
 'Black <20 yrs',
 'Asian US born',
 'Asian >20 yrs',
 'Asian <20 yrs',
 'White US born',
 'White >20 yrs',
 'White <20 yrs');

  EthnicGrDuration:array[1..NumberOfEthnicGrs] of integer=( 0, 0,20, 0, 0,20, 0, 0,20, 0, 0,20);
  NextEthnicGroup:array[1..NumberOfEthnicGrs]  of integer=( 0, 0, 2, 0, 0, 5, 0, 0, 8, 0, 0,11);
  BirthEthnicGroup:array[1..NumberOfEthnicGrs] of integer=( 1, 1, 1, 4, 4, 4, 7, 7, 7,10,10,10);

  OldNumberOfEthnicGrs = 6;
  OldEthnicGroup:array[1..NumberOfEthnicGrs]  of integer=( 3, 1, 2, 4, 1, 2, 5, 1, 2, 6, 1, 2);

NumberOfIncomeGrs = 3;
IncomeGrLabels:array[1..NumberOfIncomeGrs] of string[29]=
('Lower Income',
 'Middle Income',
 'Upper Income');
 LowIncomeDummy:array[1..NumberOfIncomeGrs] of integer = (1,0,0);
 MiddleIncomeDummy:array[1..NumberOfIncomeGrs] of integer = (0,1,0);
 HighIncomeDummy:array[1..NumberOfIncomeGrs] of integer = (0,0,1);

NumberOfWorkerGrs = 2;
WorkerGrLabels:array[1..NumberOfWorkerGrs] of string[29]=
('In Workforce',
 'Not in Workforce');

  BirthWorkerGr = 2; {new births are non-workers}

NumberOfAreaTypes = 12;
AreaTypeLabels:array[1..NumberOfAreaTypes] of string[29]=
('PHIL-Suburban',
 'PHIL-Second City',
 'PHIL-Urban',
 'PHIL-Urban Core',
 'O.PA-Rural',
 'O.PA-Suburban',
 'O.PA-Second City',
 'O.PA-Urban',
 'N.J.-Rural',
 'N.J.-Suburban',
 'N.J.-Second City',
 'N.J.-Urban');

NumberOfDensityTypes = 5;
DensityTypeLabels:array[1..NumberOfDensityTypes] of string[29]=
('Rural',
 'Suburban',
 'Second City',
 'Urban',
 'Urban Core');
AreaTypeDensity:array[1..NumberOfAreaTypes] of integer=(2,3,4,5, 1,2,3,4, 1,2,3,4);


NumberOfSubregions = 3;
SubregionLabels:array[1..NumberOfSubregions] of string[29]=
('Philadelphia',
 'Other Penn.',
 'New Jersey');
AreaTypeSubregion:array[1..NumberOfAreaTypes] of integer=(1,1,1,1, 2,2,2,2, 3,3,3,3);
(*
SubregionDensityAreaType:array[1..NumberofSubregions,1..NumberOfDensityTypes] of integer=
  ((0, 1, 2, 3, 4),
   (5, 6, 7, 8, 0),
   (9,10,11,12, 0));
*)

NumberOfMigrationTypes = 3;
MigrationTypeLabels:array[1..NumberOfMigrationTypes] of string[29]=
('Foreign Migration',
 'Domestic Migration',
 'Local Migration  ');

NumberOfEmploymentTypes = 3;
EmploymentTypeLabels:array[1..NumberOfEmploymentTypes] of string[29]=
('Retail Jobs','Service Jobs','Other Jobs');

NumberOfLandUseTypes = 4;
LandUseTypeLabels:array[1..NumberOfLandUseTypes] of string[29]=
('Non-resid. Land','Residential Land','Developable Land','Protected Land');

NumberOfRoadTypes = 3;
RoadTypeLabels:array[1..NumberOfRoadTypes] of string[29]=
('Freeways','Arterials','Local Roads');

NumberOfODTypes = 3;

NumberOfTransitTypes = 2;
TransitTypeLabels:array[1..NumberOfTransitTypes] of string[29]=
('Rail Transit','Bus Transit');

NumberOfTravelModelVariables = 56;
NumberOfTravelModelEquations = 17;
CarOwnership_CarCompetition = 1;
CarOwnership_NoCar = 2;
WorkTrip_Generation = 3;
NonWorkTrip_Generation = 4;
ChildTrip_Generation = 5;
NonWorkTrip_CarPassengerMode = 6;
NonWorkTrip_TransitMode = 7;
NonWorkTrip_WalkBikeMode = 8;
WorkTrip_CarPassengerMode = 9;
WorkTrip_TransitMode = 10;
WorkTrip_WalkBikeMode = 11;
ChildTrip_CarPassengerMode = 12;
ChildTrip_TransitMode = 13;
ChildTrip_WalkBikeMode = 14;
CarDriverTrip_Distance = 15;
CarPassengerTrip_Distance = 16;
TransitTrip_Distance = 17;

EffectCurveIntervals=20;

 {global variables}

type
TimeStepArray = array[0..NumberOfTimeSteps] of single;

AreaTypeArray = array[1..NumberOfAreaTypes] of TimeStepArray;

EmploymentArray = array
[1..NumberOfAreaTypes,
 1..NumberOfEmploymentTypes] of TimeStepArray;

LandUseArray = array
[1..NumberOfAreaTypes,
 1..NumberOfLandUseTypes] of TimeStepArray;

RoadSupplyArray = array
[1..NumberOfAreaTypes,
 1..NumberOfRoadTypes] of TimeStepArray;

TransitSupplyArray = array
[1..NumberOfAreaTypes,
 1..NumberOfTransitTypes] of TimeStepArray;

DemographicArray = array
[1..NumberOfAreaTypes,
 1..NumberOfAgeGroups,
 1..NumberOfHhldTypes,
 1..NumberOfEthnicGrs,
 1..NumberOfIncomeGrs,
 1..NumberOfWorkerGrs] of TimeStepArray;

EffectCurveArray = array[-2..EffectCurveIntervals] of single;

function EffectCurve(curvePoints:EffectCurveArray; arg:single):single;
var low,high,pointOnCurve:single; lowerPoint,higherPoint:integer;
begin
  low:=curvePoints[-2];
  high:=curvePoints[-1];
  if low>=high then begin
    writeln('Invalid endpoint arguments for effect curve ...',low:3:2,' and ',high:3:2,' Press Enter');
    EffectCurve:=curvePoints[0];
    readln;
  end else
  if (arg<= low) then EffectCurve:=curvePoints[0] else
  if (arg>=high) then EffectCurve:=curvePoints[EffectCurveIntervals] else begin
  {interpolate linearly between points}
    pointOnCurve:= ((arg-low) / (high-low)) * EffectCurveIntervals;
    lowerPoint:= trunc(pointOnCurve);
    higherPoint:=lowerPoint+1;
    EffectCurve:= curvePoints[lowerPoint]
      + (pointOnCurve-lowerPoint) * (curvePoints[higherPoint]-curvePoints[lowerPoint]);
  end;
end;


var

C_EffectOfJobDemandSupplyIndexOnEmployerAttractiveness,
C_EffectOfCommercialSpaceDemandSupplyIndexOnEmployerAttractiveness,
C_EffectOfRoadMileDemandSupplyIndexOnEmployerAttractiveness,

C_EffectOfJobDemandSupplyIndexOnResidentAttractiveness,
C_EffectOfResidentialSpaceDemandSupplyIndexOnResidentAttractiveness,
C_EffectOfRoadMileDemandSupplyIndexOnResidentAttractiveness

 :EffectCurveArray;




TravelModelParameter: array
[1..NumberOfTravelModelEquations,
 1..NumberOfTravelModelVariables] of single;

Jobs,
JobsCreated,
JobsLost,
JobsMovedOut,
JobsMovedIn : EmploymentArray;

Land,
ChangeInLandUseOut,
ChangeInLandUseIn : LandUseArray;

RoadLaneMiles,
RoadLaneMilesAdded,
RoadLaneMilesLost : RoadSupplyArray;

TransitRouteMiles,
TransitRouteMilesAdded,
TransitRouteMilesLost : TransitSupplyArray;

WorkplaceDistribution : array
[1..NumberOfAreaTypes,
 1..NumberOfAreaTypes] of TimeStepArray;

Population,
AgeingOut,
AgeingIn,
DeathsOut,
BirthsFrom,
BirthsIn,
MarriagesOut,
MarriagesIn,
DivorcesOut,
DivorcesIn,
FirstChildOut,
FirstChildIn,
EmptyNestOut,
EmptyNestIn,
LeaveNestOut,
LeaveNestIn,
WorkerStatusOut,
WorkerStatusIn,
IncomeGroupOut,
IncomeGroupIn,
AcculturationOut,
AcculturationIn,
RegionalOutmigration,
RegionalInmigration,
DomesticOutmigration,
DomesticInmigration,
ForeignOutmigration,
ForeignInmigration,
OwnCar,
ShareCar,
NoCar,
WorkTrips,
NonWorkTrips,
CarDriverWorkTrips,
CarPassengerWorkTrips,
TransitWorkTrips,
WalkBikeWorkTrips,
CarDriverWorkMiles,
CarPassengerWorkMiles,
TransitWorkMiles,
WalkBikeWorkMiles,
CarDriverNonWorkTrips,
CarPassengerNonWorkTrips,
TransitNonWorkTrips,
WalkBikeNonWorkTrips,
CarDriverNonWorkMiles,
CarPassengerNonWorkMiles,
TransitNonWorkMiles,
WalkBikeNonWorkMiles
:DemographicArray;


const
 NumberOfDemographicVariables = 47;
 DemographicVariableLabels:array[1..NumberOfDemographicVariables] of string=
 ('Population',
  'Ageing',
  'Deaths',
  'Births',
  'Marriages',
  'Divorces',
  'FirstChild',
  'EmptyNest',
  'LeaveNest',
  'ChangeStatus',
  'ChangeIncome',
  '20YearsInU',
  'AgeingIn',
  'BirthsIn',
  'MarriagesIn',
  'DivorcesIn',
  'FirstChildIn',
  'EmptyNestIn',
  'LeaveNestIn',
  'WorkforceIn',
  'IncomeGroupIn',
  '20YearsInUSIn',
  'ForeignInmigration',
  'ForeignOutmigration',
  'DomesticInmigration',
  'DomesticOutmigration',
  'RegionalInmigration',
  'RegionalOutmigration',
  'OwnCar',
  'ShareCar',
  'NoCar',
  'WorkTrips',
  'NonWorkTrips',
  'CarDriverWorkTrips',
  'CarPassengerWorkTrips',
  'TransitWorkTrips',
  'WalkBikeWorkTrips',
  'CarDriverWorkMiles',
  'CarPassengerWorkMiles',
  'TransitWorkMiles' ,
  'CarDriverNonWorkTrips',
  'CarPassengerNonWorkTrips',
  'TransitNonWorkTrips',
  'WalkBikeNonWorkTrips',
  'CarDriverNonWorkMiles',
  'CarPassengerNonWorkMiles',
  'TransitNonWorkMiles'   );

  var
 {current demographic marginals}
 AgeGroupMarginals:array[1..NumberOfDemographicVariables,0..NumberOfAgeGroups] of TimeStepArray;
 HhldTypeMarginals:array[1..NumberOfDemographicVariables,1..NumberOfHhldTypes] of TimeStepArray;
 EthnicGrMarginals:array[1..NumberOfDemographicVariables,1..NumberOfEthnicGrs] of TimeStepArray;
 IncomeGrMarginals:array[1..NumberOfDemographicVariables,1..NumberOfIncomeGrs] of TimeStepArray;
 WorkerGrMarginals:array[1..NumberOfDemographicVariables,1..NumberOfWorkerGrs] of TimeStepArray;
 AreaTypeMarginals:array[1..NumberOfDemographicVariables,1..NumberOfAreaTypes] of TimeStepArray;

 {target demographic marginals}
 AgeGroupTargetMarginals:array[1..NumberOfSubregions,1..NumberOfAgeGroups] of single;
 HhldTypeTargetMarginals:array[1..NumberOfSubregions,1..NumberOfHhldTypes] of single;
 EthnicGrTargetMarginals:array[1..NumberOfSubregions,1..NumberOfEthnicGrs] of single;
 IncomeGrTargetMarginals:array[1..NumberOfSubregions,1..NumberOfIncomeGrs] of single;
 WorkerGrTargetMarginals:array[1..NumberOfSubregions,1..NumberOfWorkerGrs] of single;
 AreaTypeTargetMarginals:array[1..NumberOfSubregions,1..NumberOfAreaTypes] of single;


BaseAverageHouseholdSize,
MigrationRateMultiplier,
BaseMortalityRate,
BaseFertilityRate,
BaseMarriageRate,
BaseDivorceRate,
BaseEmptyNestRate,
BaseLeaveNestSingleRate,
BaseLeaveNestCoupleRate,
BaseEnterWorkforceRate,
BaseLeaveWorkforceRate,
BaseEnterLowIncomeRate,
BaseLeaveLowIncomeRate,
BaseEnterHighIncomeRate,
BaseLeaveHighIncomeRate:array
[1..NumberOfAgeGroups,
 1..NumberOfHhldTypes,
 1..NumberOfEthnicGrs] of single;

MarryNoChildren_ChildrenFraction,
MarryHasChildren_ChildrenFraction,
DivorceNoChildren_ChildrenFraction,
DivorceHasChildren_ChildrenFraction,
LeaveNestSingle_ChildrenFraction,
LeaveNestCouple_ChildrenFraction: single;

BaseForeignInmigrationRate,
BaseForeignOutmigrationRate,
BaseDomesticMigrationRate,
BaseRegionalMigrationRate:single;


var
ExogenousEffectOnMortalityRate,
ExogenousEffectOnFertilityRate,
ExogenousEffectOnMarriageRate,
ExogenousEffectOnDivorceRate,
ExogenousEffectOnEmptyNestRate,
ExogenousEffectOnLeaveWorkforceRate,
ExogenousEffectOnEnterWorkforceRate,
ExogenousEffectOnLeaveLowIncomeRate,
ExogenousEffectOnEnterLowIncomeRate,
ExogenousEffectOnLeaveHighIncomeRate,
ExogenousEffectOnEnterHighIncomeRate,
ExogenousEffectOnForeignInmigrationRate,
ExogenousEffectOnForeignOutmigrationRate,
ExogenousEffectOnDomesticMigrationRate,
ExogenousEffectOnRegionalMigrationRate,
ExogenousPopulationChangeRate1,
ExogenousPopulationChangeRate2,
ExogenousPopulationChangeRate3,
ExogenousPopulationChangeRate4,
ExogenousPopulationChangeRate5,
ExogenousPopulationChangeRate6,
ExogenousPopulationChangeRate7,
ExogenousPopulationChangeRate8,
ExogenousPopulationChangeRate9,
ExogenousPopulationChangeRate10,
ExogenousPopulationChangeRate11,
ExogenousPopulationChangeRate12,
SingleNoKidsEffectOnMoveTowardsUrbanAreas,
CoupleNoKidsEffectOnMoveTowardsUrbanAreas,
SingleWiKidsEffectOnMoveTowardsUrbanAreas,
CoupleWiKidsEffectOnMoveTowardsUrbanAreas,
LowIncomeEffectOnMoveTowardsUrbanAreas,
HighIncomeEffectOnMoveTowardsUrbanAreas,
LowIncomeEffectOnMortalityRate,
HighIncomeEffectOnMortalityRate,
LowIncomeEffectOnFertilityRate,
HighIncomeEffectOnFertilityRate,
LowIncomeEffectOnMarriageRate,
HighIncomeEffectOnMarriageRate,
LowIncomeEffectOnDivorceRate,
HighIncomeEffectOnDivorceRate,
LowIncomeEffectOnEmptyNestRate,
HighIncomeEffectOnEmptyNestRate,
LowIncomeEffectOnSpacePerHousehold,
HighIncomeEffectOnSpacePerHousehold,
WorkforceChangeDelay,
IncomeChangeDelay,
ForeignInmigrationDelay,
ForeignOutmigrationDelay,
DomesticMigrationDelay,
RegionalMigrationDelay,

ExogenousEffectOnGasolinePrice,
ExogenousEffectOnSharedCarFraction,
ExogenousEffectOnNoCarFraction,
ExogenousEffectOnWorkTripRate,
ExogenousEffectOnNonworkTripRate,
ExogenousEffectOnCarPassengerModeFraction,
ExogenousEffectOnTransitModeFraction,
ExogenousEffectOnWalkBikeModeFraction,
ExogenousEffectOnCarTripDistance,
ExogenousEffectOnAgeCohortVariables,

ExogenousEffectOnJobCreationRate,
ExogenousEffectOnJobLossRate,
ExogenousEffectOnJobMoveRate,
JobCreationDelay,
JobLossDelay,
JobMoveDelay,
ExogenousEmploymentChangeRate1A,
ExogenousEmploymentChangeRate2A,
ExogenousEmploymentChangeRate3A,
ExogenousEmploymentChangeRate4A,
ExogenousEmploymentChangeRate5A,
ExogenousEmploymentChangeRate6A,
ExogenousEmploymentChangeRate7A,
ExogenousEmploymentChangeRate8A,
ExogenousEmploymentChangeRate9A,
ExogenousEmploymentChangeRate10A,
ExogenousEmploymentChangeRate11A,
ExogenousEmploymentChangeRate12A,
ExogenousEmploymentChangeRate1B,
ExogenousEmploymentChangeRate2B,
ExogenousEmploymentChangeRate3B,
ExogenousEmploymentChangeRate4B,
ExogenousEmploymentChangeRate5B,
ExogenousEmploymentChangeRate6B,
ExogenousEmploymentChangeRate7B,
ExogenousEmploymentChangeRate8B,
ExogenousEmploymentChangeRate9B,
ExogenousEmploymentChangeRate10B,
ExogenousEmploymentChangeRate11B,
ExogenousEmploymentChangeRate12B,
ExogenousEmploymentChangeRate1C,
ExogenousEmploymentChangeRate2C,
ExogenousEmploymentChangeRate3C,
ExogenousEmploymentChangeRate4C,
ExogenousEmploymentChangeRate5C,
ExogenousEmploymentChangeRate6C,
ExogenousEmploymentChangeRate7C,
ExogenousEmploymentChangeRate8C,
ExogenousEmploymentChangeRate9C,
ExogenousEmploymentChangeRate10C,
ExogenousEmploymentChangeRate11C,
ExogenousEmploymentChangeRate12C,

ExogenousEffectOnResidentialSpacePerHousehold,
ExogenousEffectOnCommercialSpacePerJob,
ExogenousEffectOnLandProtection,
ResidentialSpaceDevelopmentDelay,
ResidentialSpaceReleaseDelay,
CommercialSpaceDevelopmentDelay,
CommercialSpaceReleaseDelay,
LandProtectionProcessDelay,

ExogenousEffectOnRoadCapacityAddition,
ExogenousEffectOnTransitCapacityAddition,
ExogenousEffectOnRoadCapacityPerLane,
ExogenousEffectOnTransitCapacityPerRoute,
RoadCapacityAdditionDelay,
RoadCapacityRetirementDelay,
TransitCapacityAdditionDelay,
TransitCapacityRetirementDelay,

ExternalJobDemandSupplyIndex,
ExternalCommercialSpaceDemandSupplyIndex,
ExternalResidentialSpaceDemandSupplyIndex,
ExternalRoadMileDemandSupplyIndex

:TimeStepArray;

JobDemand,
JobSupply,
JobDemandSupplyIndex,
ResidentialSpaceDemand,
ResidentialSpaceSupply,
ResidentialSpaceDemandSupplyIndex,
CommercialSpaceDemand,
CommercialSpaceSupply,
CommercialSpaceDemandSupplyIndex,
DevelopableSpaceDemand,
DevelopableSpaceSupply,
DevelopableSpaceDemandSupplyIndex,
RoadVehicleCapacityDemandSupplyIndex,
TransitPassengerCapacityDemandSupplyIndex
:AreaTypeArray;

WorkTripRoadMileDemand,
NonWorkTripRoadMileDemand,
RoadVehicleCapacityDemand,
RoadVehicleCapacitySupply:RoadSupplyArray;

WorkTripTransitMileDemand,
NonWorkTripTransitMileDemand,
TransitPassengerCapacityDemand,
TransitPassengerCapacitySupply:TransitSupplyArray;

BaseResidentialSpacePerPerson:array[1..NumberOfAreaTypes,1..NumberOfHhldTypes] of single;
BaseCommercialSpacePerJob:array[1..NumberOfAreaTypes,1..NumberOfEmploymentTypes] of single;
BaseRoadLaneCapacityPerHour:array[1..NumberOfAreaTypes,1..NumberOfRoadTypes] of single;
BaseTransitRouteCapacityPerHour:array[1..NumberOfAreaTypes,1..NumberOfTransitTypes] of single;

FractionOfDevelopableLandAllowedForResidential:array[1..NumberOfAreaTypes] of single;
FractionOfDevelopableLandAllowedForCommercial:array[1..NumberOfAreaTypes] of single;

WeightOfJobDemandSupplyIndexInEmployerAttractiveness,
WeightOfCommercialSpaceDemandSupplyIndexInEmployerAttractiveness,
WeightOfRoadMileDemandSupplyIndexInEmployerAttractiveness
:array[1..NumberOfAreaTypes,1..NumberOfEmploymentTypes] of single;

WeightOfJobDemandSupplyIndexInResidentAttractiveness,
WeightOfResidentialSpaceDemandSupplyIndexInResidentAttractiveness,
WeightOfRoadMileDemandSupplyIndexInResidentAttractiveness
:array[1..NumberOfAreaTypes,1..NumberOfHhldTypes] of single;

WorkTripPeakHourFraction,NonWorkTripPeakHourFraction:single;

DistanceFractionByRoadType:array[1..NumberOfAreaTypes,1..NumberOfODTypes,1..NumberOfRoadTypes] of single;

WorkTripAutoVehiclePADistribution,
WorkTripAutoPersonPADistribution,
WorkTripTransitPADistribution,
NonworkTripAutoVehiclePADistribution,
NonworkTripAutoPersonPADistribution,
NonworkTripTransitPADistribution,
CarTripAverageODDistance,
TransitTripAverageODDistance,
TransitRailPAFraction:array[1..NumberOfAreaTypes,1..NumberOfAreaTypes] of single;

ODThroughDistanceFraction:array[1..NumberOfAreaTypes,1..NumberOfAreaTypes,1..NumberOfAreaTypes] of single;



Const
BaseGasolinePrice = 3.00;

var
RunLabel,InputDirectory,OutputDirectory:string;

procedure ReadUserInputData;
{const

 ScenarioUserInputsFilename = 'ScenarioUserInputs.dat';
 DemographicInitialValuesFilename  = 'DemographicInitialValues.dat';
 EmploymentInitialValuesFilename = 'EmploymentInitialValues.dat';
 LandUseInitialValuesFilename = 'LandUseInitialValues.dat';
 TransportationSupplyInitialValuesFilename = 'TransportationSupplyInitialValues.dat';
 TravelModelParameterFilename = 'TravelModelParameters.dat';
 DemographicSeedMatrixFilename = 'DemographicSeedMatrix.dat';
 DemographicTransitionRatesFilename = 'DemographicTransitionRates.dat';
}
 var {inf:text;}
    prefix:string[6]; x:string[1]; inString:string[80]; ctlFileName:string;
{indices}
Subregion,
ResAreaType,
WorkAreaType,
AreaType,
WorkerGr,
IncomeGr,
EthnicGr,
OldEthnicGr,
HhldType,
AgeGroup,
DensityType,
EmploymentType,
LandUseType,
RoadType,
TransitType,
ODType,
MigrationType,
TravelModelVariable,
TravelModelEquation,
point,
rate
: byte;
xval:single;
tempRate:array[1..15] of single;


function getExcelData(fcell,lcell:string):TDynFloat;
var datArray:TDynFloat; r,c:integer;
begin
  datArray := ReadSpreadsheetRange(worksheet,fcell,lcell);
  getExcelData:= datArray;
end;

procedure setTimeArray(var scenVar:TimeStepArray; tFirst,tLast:integer; fcell,lcell:string);
var t,ts,s,timeStepsPerValue:integer; value,previousValue:single;
begin
datArray := ReadSpreadsheetRange(worksheet,fcell,lcell);
if tFirst=tLast then begin
  value:=datArray[0,0];
  for ts:=0 to NumberOfTimeSteps do scenVar[ts]:=value;
 end else begin
  timeStepsPerValue:= round(NumberOfTimeSteps * 1.0 / (tLast-TFirst));
  ts:=0;
  for t:=tFirst to tLast do begin
    value:=datArray[0,t-tFirst];
    if t=0 then scenVar[ts]:=value else
    {do straight line interpolation between user input values for each time step}
    for s:=1 to timeStepsPerValue do begin
      ts:=ts+1;
      scenVar[ts]:=previousValue + s*1.0/timeStepsPerValue * (value - previousValue);
    end;
    previousValue:=value;
  end;
 end;
end;

var inf:text; ii,rr,cc:integer;
begin

 {read control file}
  if paramCount>0 then ctlFileName:= paramStr(1) else ctlFileName:='Baseline_Test46b.ctl';
  {write(ctlFileName); readln;}
  assign(inf,ctlFileName); reset(inf);
  repeat
    readln(inf,prefix,x,inString);
    while inString[1]=' ' do inString:= copy(inString,2,length(inString)-1);
    while inString[length(inString)]=' ' do inString:= copy(inString,1,length(inString)-1);
    for ii:=1 to length(inString) do if inString[ii]=Chr(34) then inString[ii]:=Chr(32);

    if prefix='RUNLAB' then RunLabel:=inString else
    if prefix='INPDIR' then InputDirectory:=inString else
    if prefix='OUTDIR' then OutputDirectory:=inString else
    if prefix='REGION' then Region:=StrToInt(inString) else
    if prefix='SCENAR' then Scenario:=StrToInt(inString);
  until eof(inf);
  {if InputDirectory[length(InputDirectory)]<>'\' then InputDirectory:=InputDirectory+'\';}
  if OutputDirectory[length(OutputDirectory)]<>'\' then OutputDirectory:=OutputDirectory+'\';

 workbook := TsWorkbook.Create;
 workbook.ReadFromFile(InputDirectory);

{read demographic seed matrix}
  worksheet := workbook.GetWorksheetByName('Demographic seed matrix');
  datArray := getExcelData('G6','W1805');

  rr:=0;
  for DensityType:=1 to NumberOfDensityTypes do
  for WorkerGr:=1 to NumberOfWorkerGrs do
  for IncomeGr:=1 to NumberOfIncomeGrs do
  for EthnicGr:=1 to NumberOfEthnicGrs do
  for HhldType:=1 to NumberOfHhldTypes + 1 do begin
    if HhldType<=NumberOfHhldTypes then begin {totals row from SPSS, left in for convenience}
      for AgeGroup:=1 to NumberOfAgeGroups do
      for AreaType:=1 to NumberOfAreaTypes do if AreaTypeDensity[AreaType]=DensityType then
      Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][0]:=
        datArray[rr,AgeGroup-1];
    end;
    rr:=rr+1; {next row}
  end;

{read demographic sector initial values}
  worksheet := workbook.GetWorksheetByName('Demographic initial values');

  datArray := getExcelData('B6','D22');
  for Subregion:=1 to NumberOfSubregions do
  for AgeGroup:=1 to NumberOfAgeGroups do
    AgeGroupTargetMarginals[Subregion,AgeGroup]:=datArray[AgeGroup-1,Subregion-1];

  datArray := ReadSpreadsheetRange(worksheet,'B26','D29');
  for Subregion:=1 to NumberOfSubregions do
  for HhldType:=1 to NumberOfHhldTypes do
    HhldTypeTargetMarginals[Subregion,HhldType]:=datArray[HhldType-1,Subregion-1];

  datArray := getExcelData('B33','D44');
  for Subregion:=1 to NumberOfSubregions do
  for EthnicGr:=1 to NumberOfEthnicGrs do
    EthnicGrTargetMarginals[Subregion,EthnicGr]:=datArray[EthnicGr-1,Subregion-1];

  datArray := getExcelData('B48','D50');
  for Subregion:=1 to NumberOfSubregions do
  for IncomeGr:=1 to NumberOfIncomeGrs do
    IncomeGrTargetMarginals[Subregion,IncomeGr]:=datArray[IncomeGr-1,Subregion-1];

  datArray := getExcelData('B54','D55');
  for Subregion:=1 to NumberOfSubregions do
  for WorkerGr:=1 to NumberOfWorkerGrs do
    WorkerGrTargetMarginals[Subregion,WorkerGr]:=datArray[WorkerGr-1,Subregion-1];

  datArray := getExcelData('B59','D70');
  for Subregion:=1 to NumberOfSubregions do
  for AreaType:=1 to NumberOfAreaTypes do
    AreaTypeTargetMarginals[Subregion,AreaType]:=datArray[AreaType-1,Subregion-1];


{read employment sector initial values}
  worksheet := workbook.GetWorksheetByName('Employment initial values');

  datArray := getExcelData('B5','D16');
  for AreaType:=1 to NumberOfAreaTypes do
  for EmploymentType:=1 to NumberOfEmploymentTypes do begin
      Jobs[AreaType][EmploymentType][0]:=datArray[AreaType-1,EmploymentType-1];
  end;

  setTimeArray(JobCreationDelay,0,0,'B26','B26');
  setTimeArray(JobLossDelay,0,0,'B27','B27');
  setTimeArray(JobMoveDelay,0,0,'B28','B28');


  datArray := getExcelData('B33','D47');
  rr:=0;
  for DensityType:=1 to NumberOfDensityTypes do
  for EmploymentType:=1 to NumberOfEmploymentTypes do begin
    for AreaType:=1 to NumberOfAreaTypes do if AreaTypeDensity[AreaType]=DensityType then
      WeightOfJobDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]:=datArray[rr,0];
    for AreaType:=1 to NumberOfAreaTypes do if AreaTypeDensity[AreaType]=DensityType then
      WeightOfCommercialSpaceDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]:=datArray[rr,1];
    for AreaType:=1 to NumberOfAreaTypes do if AreaTypeDensity[AreaType]=DensityType then
      WeightOfRoadMileDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]:=datArray[rr,2];
    rr:=rr+1;
  end;

  datArray := getExcelData('A51','D71');
  for point:=0 to EffectCurveIntervals do begin
    xval:=datArray[point,0];
    C_EffectOfJobDemandSupplyIndexOnEmployerAttractiveness[point]:=datArray[point,1];
    C_EffectOfCommercialSpaceDemandSupplyIndexOnEmployerAttractiveness[point]:=datArray[point,2];
    C_EffectOfRoadMileDemandSupplyIndexOnEmployerAttractiveness[point]:=datArray[point,3];
    if point=0 then begin
       C_EffectOfJobDemandSupplyIndexOnEmployerAttractiveness[-2]:=xval;
       C_EffectOfCommercialSpaceDemandSupplyIndexOnEmployerAttractiveness[-2]:=xval;
       C_EffectOfRoadMileDemandSupplyIndexOnEmployerAttractiveness[-2]:=xval;
     end else
     if point=EffectCurveIntervals then begin
       C_EffectOfJobDemandSupplyIndexOnEmployerAttractiveness[-1]:=xval;
       C_EffectOfCommercialSpaceDemandSupplyIndexOnEmployerAttractiveness[-1]:=xval;
       C_EffectOfRoadMileDemandSupplyIndexOnEmployerAttractiveness[-1]:=xval;
     end;
  end;

{read land use sector initial values}
  worksheet := workbook.GetWorksheetByName('Land use initial values');

  datArray := getExcelData('B4','E15');
  for AreaType:=1 to NumberOfAreaTypes do
  for LandUseType:=1 to NumberOfLandUseTypes do begin
    Land[AreaType][LandUseType][0]:=datArray[AreaType-1,LandUseType-1];
  end;

  setTimeArray(ResidentialSpaceDevelopmentDelay,0,0,'B18','B18');
  setTimeArray(ResidentialSpaceReleaseDelay,0,0,'B19','B19');
  setTimeArray(CommercialSpaceDevelopmentDelay,0,0,'B20','B20');
  setTimeArray(CommercialSpaceReleaseDelay,0,0,'B21','B21');
  setTimeArray(LandProtectionProcessDelay,0,0,'B22','B22');

  datArray := getExcelData('B26','E37');
  for AreaType:=1 to NumberOfAreaTypes do begin
    for HhldType:=1 to NumberOfHhldTypes do
    BaseResidentialSpacePerPerson[AreaType][HhldType]:=datArray[AreaType-1,HhldType-1];
  end;

  datArray := getExcelData('B41','D52');
  for AreaType:=1 to NumberOfAreaTypes do begin
    for EmploymentType:=1 to NumberOfEmploymentTypes do
    BaseCommercialSpacePerJob[AreaType][EmploymentType]:=datArray[AreaType-1,EmploymentType-1];
  end;

  datArray := getExcelData('B56','C67');
  for AreaType:=1 to NumberOfAreaTypes do begin
    FractionOfDevelopableLandAllowedForCommercial[AreaType]:=datArray[AreaType-1,0];
    FractionOfDevelopableLandAllowedForResidential[AreaType]:=datArray[AreaType-1,1];
  end;

{read transportation supply sector initial values}
  worksheet := workbook.GetWorksheetByName('Transportation initial values');

  datArray := getExcelData('B4','F15');
  for AreaType:=1 to NumberOfAreaTypes do begin
    for RoadType:=1 to NumberOfRoadTypes do
      RoadLaneMiles[AreaType][RoadType][0]:=datArray[AreaType-1,RoadType-1];
    for TransitType:=1 to NumberOfTransitTypes do
      TransitRouteMiles[AreaType][TransitType][0]:=datArray[AreaType-1,NumberOfRoadTypes+TransitType-1];
  end;

  setTimeArray(RoadCapacityAdditionDelay,0,0,'B18','B18');
  setTimeArray(RoadCapacityRetirementDelay,0,0,'B19','B19');
  setTimeArray(TransitCapacityAdditionDelay,0,0,'B20','B20');
  setTimeArray(TransitCapacityRetirementDelay,0,0,'B21','B21');


  datArray := getExcelData('B25','D36');
  for AreaType:=1 to NumberOfAreaTypes do begin
    for RoadType:=1 to NumberOfRoadTypes do
      BaseRoadLaneCapacityPerHour[AreaType][RoadType]:=datArray[AreaType-1,RoadType-1];
  end;

  datArray := getExcelData('B40','C51');
  for AreaType:=1 to NumberOfAreaTypes do begin
    for TransitType:=1 to NumberOfTransitTypes do
      BaseTransitRouteCapacityPerHour[AreaType][TransitType]:=datArray[AreaType-1,TransitType-1];
  end;

  datArray := getExcelData('B54','B55');
  WorkTripPeakHourFraction:=datArray[0,0];
  NonWorkTripPeakHourFraction:=datArray[1,0];

  datArray := getExcelData('B58','J69');
  for AreaType:=1 to NumberOfAreaTypes do begin
    cc:=0;
    for ODType:=1 to NumberOfODTypes do
    for RoadType:=1 to NumberOfRoadTypes do begin
      DistanceFractionByRoadType[AreaType][ODType][RoadType]:=datArray[AreaType-1,cc];
      cc:=cc+1;
    end;
  end;

  datArray := getExcelData('B72','M83');
  for ResAreaType:=1 to NumberOfAreaTypes do
  for AreaType:=1 to NumberOfAreaTypes do begin
    WorkTripAutoVehiclePADistribution[ResAreaType][AreaType]:=datArray[ResAreaType-1,AreaType-1];
  end;

  datArray := getExcelData('B86','M97');
  for ResAreaType:=1 to NumberOfAreaTypes do
  for AreaType:=1 to NumberOfAreaTypes do begin
    WorkTripAutoPersonPADistribution[ResAreaType][AreaType]:=datArray[ResAreaType-1,AreaType-1];
    WorkplaceDistribution[ResAreaType][AreaType][0]:=WorkTripAutoPersonPADistribution[ResAreaType][AreaType];
  end;

  datArray := getExcelData('B100','M111');
  for ResAreaType:=1 to NumberOfAreaTypes do
  for AreaType:=1 to NumberOfAreaTypes do begin
    WorkTripTransitPADistribution[ResAreaType][AreaType]:=datArray[ResAreaType-1,AreaType-1];
  end;

  datArray := getExcelData('B114','M125');
  for ResAreaType:=1 to NumberOfAreaTypes do
  for AreaType:=1 to NumberOfAreaTypes do begin
    NonworkTripAutoVehiclePADistribution[ResAreaType][AreaType]:=datArray[ResAreaType-1,AreaType-1];
  end;

  datArray := getExcelData('B128','M139');
  for ResAreaType:=1 to NumberOfAreaTypes do
  for AreaType:=1 to NumberOfAreaTypes do begin
    NonworkTripAutoPersonPADistribution[ResAreaType][AreaType]:=datArray[ResAreaType-1,AreaType-1];
  end;

  datArray := getExcelData('B142','M153');
  for ResAreaType:=1 to NumberOfAreaTypes do
  for AreaType:=1 to NumberOfAreaTypes do begin
    NonworkTripTransitPADistribution[ResAreaType][AreaType]:=datArray[ResAreaType-1,AreaType-1];
  end;

  datArray := getExcelData('B156','M167');
  for ResAreaType:=1 to NumberOfAreaTypes do
  for AreaType:=1 to NumberOfAreaTypes do begin
    CarTripAverageODDistance[ResAreaType][AreaType]:=datArray[ResAreaType-1,AreaType-1];
  end;

  datArray := getExcelData('B170','M181');
  for ResAreaType:=1 to NumberOfAreaTypes do
  for AreaType:=1 to NumberOfAreaTypes do begin
    TransitTripAverageODDistance[ResAreaType][AreaType]:=datArray[ResAreaType-1,AreaType-1];
  end;

  datArray := getExcelData('B184','M195');
  for ResAreaType:=1 to NumberOfAreaTypes do
  for AreaType:=1 to NumberOfAreaTypes do begin
    TransitRailPAFraction[ResAreaType][AreaType]:=datArray[ResAreaType-1,AreaType-1];
  end;

  datArray := getExcelData('B198','M275');
  rr:=0;
  for ResAreaType:=1 to NumberOfAreaTypes do
  for WorkAreaType:=ResAreaType to NumberOfAreaTypes do begin
    for AreaType:=1 to NumberOfAreaTypes do begin
      ODThroughDistanceFraction[ResAreaType][WorkAreaType][AreaType]:=datArray[rr,AreaType-1];
      ODThroughDistanceFraction[WorkAreaType][ResAreaType][AreaType]:=datArray[rr,AreaType-1];
    end;
    rr:=rr+1;
  end;

  {read demographic base demographic transition rates}
  worksheet := workbook.GetWorksheetByName('Demographic transition rates');

  datArray := getExcelData('D5','R412');

  rr:=0;
  for AgeGroup:=1 to NumberOfAgeGroups do
  for OldEthnicGr:=1 to OldNumberOfEthnicGrs do
  for HhldType:=1 to NumberOfHhldTypes do begin
     for rate:=1 to 15 do tempRate[rate]:=datArray[rr,rate-1];
     rr:=rr+1;

     for EthnicGr:=1 to NumberOfEthnicGrs do if OldEthnicGroup[EthnicGr]=OldEthnicGr then begin
       BaseAverageHouseholdSize[AgeGroup][HHldType][EthnicGr]:=tempRate[1];
       BaseMortalityRate[AgeGroup][HHldType][EthnicGr]:=tempRate[2];
       BaseFertilityRate[AgeGroup][HHldType][EthnicGr]:=tempRate[3];
       BaseMarriageRate[AgeGroup][HHldType][EthnicGr]:=tempRate[4];
       BaseDivorceRate[AgeGroup][HHldType][EthnicGr]:=tempRate[5];
       BaseLeaveNestSingleRate[AgeGroup][HHldType][EthnicGr]:=tempRate[6];
       BaseLeaveNestCoupleRate[AgeGroup][HHldType][EthnicGr]:=tempRate[7];
       BaseEmptyNestRate[AgeGroup][HHldType][EthnicGr]:=tempRate[8];
       BaseEnterLowIncomeRate[AgeGroup][HHldType][EthnicGr]:=tempRate[9];
       BaseLeaveLowIncomeRate[AgeGroup][HHldType][EthnicGr]:=tempRate[10];
       BaseEnterHighIncomeRate[AgeGroup][HHldType][EthnicGr]:=tempRate[11];
       BaseLeaveHighIncomeRate[AgeGroup][HHldType][EthnicGr]:=tempRate[12];
       BaseEnterWorkforceRate[AgeGroup][HHldType][EthnicGr]:=tempRate[13];
       BaseLeaveWorkforceRate[AgeGroup][HHldType][EthnicGr]:=tempRate[14];
       MigrationRateMultiplier[AgeGroup][HHldType][EthnicGr]:=tempRate[15];
     end;
  end;

  setTimeArray(WorkforceChangeDelay,0,0,'B415','B415');
  setTimeArray(IncomeChangeDelay,0,0,'B416','B416');
  setTimeArray(ForeignInmigrationDelay,0,0,'B417','B417');
  setTimeArray(ForeignOutmigrationDelay,0,0,'B418','B418');
  setTimeArray(DomesticMigrationDelay,0,0,'B419','B419');
  setTimeArray(RegionalMigrationDelay,0,0,'B420','B420');

  datArray := getExcelData('D424','D429');
  MarryNoChildren_ChildrenFraction:=datArray[0,0];
  MarryHasChildren_ChildrenFraction:=datArray[1,0];
  DivorceNoChildren_ChildrenFraction:=datArray[2,0];
  DivorceHasChildren_ChildrenFraction:=datArray[3,0];
  LeaveNestSingle_ChildrenFraction:=datArray[4,0];
  LeaveNestCouple_ChildrenFraction:=datArray[5,0];

  datArray := getExcelData('B432','B435');
  BaseForeignInmigrationRate:=datArray[0,0];
  BaseForeignOutmigrationRate:=datArray[1,0];
  BaseDomesticMigrationRate:=datArray[2,0];
  BaseRegionalMigrationRate:=datArray[3,0];

  datArray := getExcelData('B441','D455');
  rr:=0;
  for DensityType:=1 to NumberOfDensityTypes do
  for MigrationType:=1 to NumberOfMigrationTypes do begin
    for AreaType:=1 to NumberOfAreaTypes do if AreaTypeDensity[AreaType]=DensityType then
      WeightOfJobDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]:=datArray[rr,MigrationType-1];
    for AreaType:=1 to NumberOfAreaTypes do if AreaTypeDensity[AreaType]=DensityType then
      WeightOfResidentialSpaceDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]:=datArray[rr,MigrationType-1];
    for AreaType:=1 to NumberOfAreaTypes do if AreaTypeDensity[AreaType]=DensityType then
      WeightOfRoadMileDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]:=datArray[rr,MigrationType-1];
    rr:=rr+1;
  end;

  datArray := getExcelData('A459','D479');
  for point:=0 to EffectCurveIntervals do begin
     xval:=datArray[point,0];
     C_EffectOfJobDemandSupplyIndexOnResidentAttractiveness[point]:=datArray[point,1];
     C_EffectOfResidentialSpaceDemandSupplyIndexOnResidentAttractiveness[point]:=datArray[point,2];
     C_EffectOfRoadMileDemandSupplyIndexOnResidentAttractiveness[point]:=datArray[point,3];
     if point=0 then begin
         C_EffectOfJobDemandSupplyIndexOnResidentAttractiveness[-2]:=xval;
         C_EffectOfResidentialSpaceDemandSupplyIndexOnResidentAttractiveness[-2]:=xval;
         C_EffectOfRoadMileDemandSupplyIndexOnResidentAttractiveness[-2]:=xval;
     end else
     if point=EffectCurveIntervals then begin
         C_EffectOfJobDemandSupplyIndexOnResidentAttractiveness[-1]:=xval;
         C_EffectOfResidentialSpaceDemandSupplyIndexOnResidentAttractiveness[-1]:=xval;
         C_EffectOfRoadMileDemandSupplyIndexOnResidentAttractiveness[-1]:=xval;
     end;
  end;

{ read travel demand model parameters}
  worksheet := workbook.GetWorksheetByName('Travel behavior models');
  datArray := getExcelData('B6','R61');

  for TravelModelVariable:=1 to NumberOfTravelModelVariables do begin
    for TravelModelEquation:=1 to NumberOfTravelModelEquations do
      TravelModelParameter[TravelModelEquation][TravelModelVariable]:=
        datArray[TravelModelVariable-1,TravelModelEquation-1];
  end;


{read Exogenous user inputs}

  worksheet := workbook.GetWorksheetByIndex(Scenario);

{base year}
  datArray := getExcelData('B2','B2');
  Year:= datArray[0,0];

{demographic sector}

  setTimeArray(ExogenousEffectOnMortalityRate,0,10,         'B4','L4');
  setTimeArray(ExogenousEffectOnFertilityRate,0,10,         'B5','L5');
  setTimeArray(ExogenousEffectOnMarriageRate,0,10,          'B6','L6');
  setTimeArray(ExogenousEffectOnDivorceRate,0,10,           'B7','L7');
  setTimeArray(ExogenousEffectOnEmptyNestRate,0,10,         'B8','L8');
  setTimeArray(ExogenousEffectOnLeaveWorkforceRate,0,10,    'B9','L9');
  setTimeArray(ExogenousEffectOnEnterWorkforceRate,0,10,    'B10','L10');
  setTimeArray(ExogenousEffectOnLeaveLowIncomeRate,0,10,    'B11','L11');
  setTimeArray(ExogenousEffectOnEnterLowIncomeRate,0,10,    'B12','L12');
  setTimeArray(ExogenousEffectOnLeaveHighIncomeRate,0,10,   'B13','L13');
  setTimeArray(ExogenousEffectOnEnterHighIncomeRate,0,10,   'B14','L14');
  setTimeArray(ExogenousEffectOnForeignInmigrationRate,0,10,'B15','L15');
  setTimeArray(ExogenousEffectOnForeignOutmigrationRate,0,10,'B16','L16');
  setTimeArray(ExogenousEffectOnDomesticMigrationRate,0,10, 'B17','L17');
  setTimeArray(ExogenousEffectOnRegionalMigrationRate,0,10, 'B18','L18');
  setTimeArray(ExogenousPopulationChangeRate1,0,10,         'B19','L19');
  setTimeArray(ExogenousPopulationChangeRate2,0,10,         'B20','L20');
  setTimeArray(ExogenousPopulationChangeRate3,0,10,         'B21','L21');
  setTimeArray(ExogenousPopulationChangeRate4,0,10,         'B22','L22');
  setTimeArray(ExogenousPopulationChangeRate5,0,10,         'B23','L23');
  setTimeArray(ExogenousPopulationChangeRate6,0,10,         'B24','L24');
  setTimeArray(ExogenousPopulationChangeRate7,0,10,         'B25','L25');
  setTimeArray(ExogenousPopulationChangeRate8,0,10,         'B26','L26');
  setTimeArray(ExogenousPopulationChangeRate9,0,10,         'B27','L27');
  setTimeArray(ExogenousPopulationChangeRate10,0,10,        'B28','L28');
  setTimeArray(ExogenousPopulationChangeRate11,0,10,        'B29','L29');
  setTimeArray(ExogenousPopulationChangeRate12,0,10,        'B30','L30');
  setTimeArray(SingleNoKidsEffectOnMoveTowardsUrbanAreas,0,10, 'B31','L31');
  setTimeArray(CoupleNoKidsEffectOnMoveTowardsUrbanAreas,0,10, 'B32','L32');
  setTimeArray(SingleWiKidsEffectOnMoveTowardsUrbanAreas,0,10, 'B33','L33');
  setTimeArray(CoupleWiKidsEffectOnMoveTowardsUrbanAreas,0,10, 'B34','L34');

  setTimeArray(LowIncomeEffectOnMoveTowardsUrbanAreas,0,10, 'B35','L35');
  setTimeArray(HighIncomeEffectOnMoveTowardsUrbanAreas,0,10,'B36','L36');
  setTimeArray(LowIncomeEffectOnMortalityRate,0,10,         'B37','L37');
  setTimeArray(HighIncomeEffectOnMortalityRate,0,10,        'B38','L38');
  setTimeArray(LowIncomeEffectOnFertilityRate,0,10,         'B39','L39');
  setTimeArray(HighIncomeEffectOnFertilityRate,0,10,        'B40','L40');
  setTimeArray(LowIncomeEffectOnMarriageRate,0,10,          'B41','L41');
  setTimeArray(HighIncomeEffectOnMarriageRate,0,10,         'B42','L42');
  setTimeArray(LowIncomeEffectOnDivorceRate,0,10,           'B43','L43');
  setTimeArray(HighIncomeEffectOnDivorceRate,0,10,          'B44','L44');
  setTimeArray(LowIncomeEffectOnEmptyNestRate,0,10,         'B45','L45');
  setTimeArray(HighIncomeEffectOnEmptyNestRate,0,10,        'B46','L46');
  setTimeArray(LowIncomeEffectOnSpacePerHousehold,0,10,     'B47','L47');
  setTimeArray(HighIncomeEffectOnSpacePerHousehold,0,10,    'B48','L48');
  {travel behavior subsector}
  setTimeArray(ExogenousEffectOnGasolinePrice,0,10,         'B51','L51');
  setTimeArray(ExogenousEffectOnSharedCarFraction,0,10,     'B52','L52');
  setTimeArray(ExogenousEffectOnNoCarFraction,0,10,         'B53','L53');
  setTimeArray(ExogenousEffectOnWorkTripRate,0,10,          'B54','L54');
  setTimeArray(ExogenousEffectOnNonworkTripRate,0,10,       'B55','L55');
  setTimeArray(ExogenousEffectOnCarPassengerModeFraction,0,10,'B56','L56');
  setTimeArray(ExogenousEffectOnTransitModeFraction,0,10,   'B57','L57');
  setTimeArray(ExogenousEffectOnWalkBikeModeFraction,0,10,  'B58','L58');
  setTimeArray(ExogenousEffectOnCarTripDistance,0,10,       'B59','L59');
  setTimeArray(ExogenousEffectOnAgeCohortVariables,0,10,    'B60','L60');
  {employment sector}
  setTimeArray(ExogenousEffectOnJobCreationRate,0,10,       'B63','L63');
  setTimeArray(ExogenousEffectOnJobLossRate,0,10,           'B64','L64');
  setTimeArray(ExogenousEffectOnJobMoveRate,0,10,           'B65','L65');
  setTimeArray(ExogenousEmploymentChangeRate1A,0,10,         'B66','L66');
  setTimeArray(ExogenousEmploymentChangeRate1B,0,10,         'B67','L67');
  setTimeArray(ExogenousEmploymentChangeRate1C,0,10,         'B68','L68');
  setTimeArray(ExogenousEmploymentChangeRate2A,0,10,         'B69','L69');
  setTimeArray(ExogenousEmploymentChangeRate2B,0,10,         'B70','L70');
  setTimeArray(ExogenousEmploymentChangeRate2C,0,10,         'B71','L71');
  setTimeArray(ExogenousEmploymentChangeRate3A,0,10,         'B72','L72');
  setTimeArray(ExogenousEmploymentChangeRate3B,0,10,         'B73','L73');
  setTimeArray(ExogenousEmploymentChangeRate3C,0,10,         'B74','L74');
  setTimeArray(ExogenousEmploymentChangeRate4A,0,10,         'B75','L75');
  setTimeArray(ExogenousEmploymentChangeRate4B,0,10,         'B76','L76');
  setTimeArray(ExogenousEmploymentChangeRate4C,0,10,         'B77','L77');
  setTimeArray(ExogenousEmploymentChangeRate5A,0,10,         'B78','L78');
  setTimeArray(ExogenousEmploymentChangeRate5B,0,10,         'B79','L79');
  setTimeArray(ExogenousEmploymentChangeRate5C,0,10,         'B80','L80');
  setTimeArray(ExogenousEmploymentChangeRate6A,0,10,         'B81','L81');
  setTimeArray(ExogenousEmploymentChangeRate6B,0,10,         'B82','L82');
  setTimeArray(ExogenousEmploymentChangeRate6C,0,10,         'B83','L83');
  setTimeArray(ExogenousEmploymentChangeRate7A,0,10,         'B84','L84');
  setTimeArray(ExogenousEmploymentChangeRate7B,0,10,         'B85','L85');
  setTimeArray(ExogenousEmploymentChangeRate7C,0,10,         'B86','L86');
  setTimeArray(ExogenousEmploymentChangeRate8A,0,10,         'B87','L87');
  setTimeArray(ExogenousEmploymentChangeRate8B,0,10,         'B88','L88');
  setTimeArray(ExogenousEmploymentChangeRate8C,0,10,         'B89','L89');
  setTimeArray(ExogenousEmploymentChangeRate9A,0,10,         'B90','L90');
  setTimeArray(ExogenousEmploymentChangeRate9B,0,10,         'B91','L91');
  setTimeArray(ExogenousEmploymentChangeRate9C,0,10,         'B92','L92');
  setTimeArray(ExogenousEmploymentChangeRate10A,0,10,        'B93','L93');
  setTimeArray(ExogenousEmploymentChangeRate10B,0,10,        'B94','L94');
  setTimeArray(ExogenousEmploymentChangeRate10C,0,10,        'B95','L95');
  setTimeArray(ExogenousEmploymentChangeRate11A,0,10,        'B96','L96');
  setTimeArray(ExogenousEmploymentChangeRate11B,0,10,        'B97','L97');
  setTimeArray(ExogenousEmploymentChangeRate11C,0,10,        'B98','L98');
  setTimeArray(ExogenousEmploymentChangeRate12A,0,10,        'B99','L99');
  setTimeArray(ExogenousEmploymentChangeRate12B,0,10,        'B100','L100');
  setTimeArray(ExogenousEmploymentChangeRate12C,0,10,        'B101','L101');
  {land use sector}
  setTimeArray(ExogenousEffectOnResidentialSpacePerHousehold,0,10,'B104','L104');
  setTimeArray(ExogenousEffectOnCommercialSpacePerJob,0,10, 'B105','L105');
  setTimeArray(ExogenousEffectOnLandProtection,0,10,        'B106','L106');
  {transport supply sector}
  setTimeArray(ExogenousEffectOnRoadCapacityAddition,0,10,   'B109','L109');
  setTimeArray(ExogenousEffectOnTransitCapacityAddition,0,10,'B110','L110');
  setTimeArray(ExogenousEffectOnRoadCapacityPerLane,0,10,    'B111','L111');
  setTimeArray(ExogenousEffectOnTransitCapacityPerRoute,0,10,'B112','L112');
  {external indices}
  setTimeArray(ExternalJobDemandSupplyIndex,0,10,             'B115','L115');
  setTimeArray(ExternalCommercialSpaceDemandSupplyIndex,0,10, 'B116','L116');
  setTimeArray(ExternalResidentialSpaceDemandSupplyIndex,0,10,'B117','L117');
  setTimeArray(ExternalRoadMileDemandSupplyIndex,0,10,        'B118','L118');

end;

procedure CalculateDemographicMarginals (Demvar:integer; timeStep:integer);
var cellValue:single;
{indices}
AreaType,
WorkerGr,
IncomeGr,
EthnicGr,
HhldType,
AgeGroup,
LoIndex,HiIndex,DemIndex: byte;
{subprocedure to recalculate the marginals}
begin

  if Demvar=0 then LoIndex:=1 else LoIndex:=Demvar;
  if Demvar=0 then HiIndex:=NumberOfSubregions else HiIndex:=Demvar;

  {empty the marginals}
  for DemIndex:=LoIndex to HiIndex do begin
    for AgeGroup:=0 to NumberOfAgeGroups do AgeGroupMarginals[DemIndex][AgeGroup][timeStep]:=0;
    for HhldType:=1 to NumberOfHhldTypes do HhldTypeMarginals[DemIndex][HhldType][timeStep]:=0;
    for EthnicGr:=1 to NumberOfEthnicGrs do EthnicGrMarginals[DemIndex][EthnicGr][timeStep]:=0;
    for IncomeGr:=1 to NumberOfIncomeGrs do IncomeGrMarginals[DemIndex][IncomeGr][timeStep]:=0;
    for WorkerGr:=1 to NumberOfWorkerGrs do WorkerGrMarginals[DemIndex][WorkerGr][timeStep]:=0;
    for AreaType:=1 to NumberOfAreaTypes do AreaTypeMarginals[DemIndex][AreaType][timeStep]:=0;
  end;
  {loop on all cells and accumulate marginals}
  for AreaType:=1 to NumberOfAreaTypes do
  for WorkerGr:=1 to NumberOfWorkerGrs do
  for IncomeGr:=1 to NumberOfIncomeGrs do
  for EthnicGr:=1 to NumberOfEthnicGrs do
  for HhldType:=1 to NumberOfHhldTypes do
  for AgeGroup:=1 to NumberOfAgeGroups do begin
    if Demvar<=1 then cellValue:=Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=2 then cellValue:=AgeingOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=3 then cellValue:=DeathsOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=4 then cellValue:=BirthsFrom[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=5 then cellValue:=MarriagesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=6 then cellValue:=DivorcesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=7 then cellValue:=FirstChildOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=8 then cellValue:=EmptyNestOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=9 then cellValue:=LeaveNestOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=10 then cellValue:=WorkerStatusOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=11 then cellValue:=IncomeGroupOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=12 then cellValue:=AcculturationOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=13 then cellValue:=AgeingIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=14 then cellValue:=BirthsIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=15 then cellValue:=MarriagesIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=16 then cellValue:=DivorcesIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=17 then cellValue:=FirstChildIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=18 then cellValue:=EmptyNestIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=19 then cellValue:=LeaveNestIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=20 then cellValue:=WorkerStatusIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=21 then cellValue:=IncomeGroupIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=22 then cellValue:=AcculturationIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep];
    if Demvar=23 then cellValue:=ForeignInmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=24 then cellValue:=ForeignOutmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=25 then cellValue:=DomesticInmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=26 then cellValue:=DomesticOutmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=27 then cellValue:=RegionalInmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=28 then cellValue:=RegionalOutmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=29 then cellValue:=OwnCar[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=30 then cellValue:=ShareCar[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=31 then cellValue:=NoCar[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=32 then cellValue:=WorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=33 then cellValue:=NonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=34 then cellValue:=CarDriverWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=35 then cellValue:=CarPassengerWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=36 then cellValue:=TransitWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=37 then cellValue:=WalkBikeWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=38 then cellValue:=CarDriverWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=39 then cellValue:=CarPassengerWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=40 then cellValue:=TransitWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=41 then cellValue:=CarDriverNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=42 then cellValue:=CarPassengerNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=43 then cellValue:=TransitNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=44 then cellValue:=WalkBikeNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=45 then cellValue:=CarDriverNonWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=46 then cellValue:=CarPassengerNonWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    if Demvar=47 then cellValue:=TransitNonWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] else
    begin end;

    if Demvar=0 then DemIndex:=AreaTypeSubregion[areaType] else DemIndex:=Demvar;
    AgeGroupMarginals[DemIndex][   0    ][timeStep]:=AgeGroupMarginals[DemIndex][   0    ][timeStep] + cellValue;
    AgeGroupMarginals[DemIndex][AgeGroup][timeStep]:=AgeGroupMarginals[DemIndex][AgeGroup][timeStep] + cellValue;
    HhldTypeMarginals[DemIndex][HhldType][timeStep]:=HhldTypeMarginals[DemIndex][HhldType][timeStep] + cellValue;
    EthnicGrMarginals[DemIndex][EthnicGr][timeStep]:=EthnicGrMarginals[DemIndex][EthnicGr][timeStep] + cellValue;
    IncomeGrMarginals[DemIndex][IncomeGr][timeStep]:=IncomeGrMarginals[DemIndex][IncomeGr][timeStep] + cellValue;
    WorkerGrMarginals[DemIndex][WorkerGr][timeStep]:=WorkerGrMarginals[DemIndex][WorkerGr][timeStep] + cellValue;
    AreaTypeMarginals[DemIndex][AreaType][timeStep]:=AreaTypeMarginals[DemIndex][AreaType][timeStep] + cellValue;
  end; {cells}

end; {CalculateDemographicMarginals}

{Procedure to initialize the Population for the region}
procedure InitializePopulation;
const
 IPFIterations = 15;
var
 demVar,iteration,dimension:integer;
 current,target:double;

{indices}
Subregion,
AreaType,
WorkerGr,
IncomeGr,
EthnicGr,
HhldType,
AgeGroup: byte;


begin {InitializePopulation}

  demVar := 0;  {population by subregion}

  {perform IPF to get the marginals to match the trarget marginals for the region}
  {perform the specified number of iterations}
  for iteration:=1 to IPFIterations do begin
    {loop on each marginal dimension}
    for dimension:=1 to NumberOfDemographicDimensions do begin
      {(re)calculate the current population marginals}
      CalculateDemographicMarginals(demVar,0);

      {loop on all the cells and adjust the current cell values to match the target marginal on the dimension}
      for AreaType:=1 to NumberOfAreaTypes do
      for WorkerGr:=1 to NumberOfWorkerGrs do
      for IncomeGr:=1 to NumberOfIncomeGrs do
      for EthnicGr:=1 to NumberOfEthnicGrs do
      for HhldType:=1 to NumberOfHhldTypes do
      for AgeGroup:=1 to NumberOfAgeGroups do begin

        Subregion:=AreaTypeSubregion[areaType];
        if dimension=1 then begin current:=AreaTypeMarginals[Subregion][AreaType][0]; target:=AreaTypeTargetMarginals[Subregion][AreaType]; end else
        if dimension=2 then begin current:=AgeGroupMarginals[Subregion][AgeGroup][0]; target:=AgeGroupTargetMarginals[Subregion][AgeGroup]; end else
        if dimension=3 then begin current:=HhldTypeMarginals[Subregion][HhldType][0]; target:=HhldTypeTargetMarginals[Subregion][HhldType]; end else
        if dimension=4 then begin current:=EthnicGrMarginals[Subregion][EthnicGr][0]; target:=EthnicGrTargetMarginals[Subregion][EthnicGr]; end else
        if dimension=5 then begin current:=IncomeGrMarginals[Subregion][IncomeGr][0]; target:=IncomeGrTargetMarginals[Subregion][IncomeGr]; end else
        if dimension=6 then begin current:=WorkerGrMarginals[Subregion][WorkerGr][0]; target:=WorkerGrTargetMarginals[Subregion][WorkerGr]; end;

        if current>0 then begin
           Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][0]:=
           Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][0] * target/current;
           {writeln(iteration:2,dimension:2,target:10:0,current:10:0,target/current:8:3);}
        end;
      end; {cells}
      {readln;}
    end; {dimensions}
  end; {iterations}

  demVar:=1;
  CalculateDemographicMarginals(demVar,0);

end; {InitializePopulation}


procedure CalculateDemographicFeedbacks(timeStep:integer);
var
{indices}

WorkAreaType,
DestAreaType,
RoadAreaType,
ODType,
RoadType,
TransitType,
AreaType,
WorkerGr,
IncomeGr,
EthnicGr,
HhldType,
AgeGroup: byte;
SpacePerPerson, residents, commuters, autotrips, transittrips, tdistance, rdistance: single;
begin

  for AreaType:=1 to NumberOfAreaTypes do begin
      JobDemand[AreaType][timeStep]:=0;
      ResidentialSpaceDemand[AreaType][timeStep]:=0;
      for RoadType:=1 to NumberOfRoadTypes do begin
        WorkTripRoadMileDemand[AreaType][RoadType][timeStep]:=0;
        NonWorkTripRoadMileDemand[AreaType][RoadType][timeStep]:=0;
      end;
      for TransitType:=1 to NumberOfTransitTypes do begin
        WorkTripTransitMileDemand[AreaType][TransitType][timeStep]:=0;
        NonWorkTripTransitMileDemand[AreaType][TransitType][timeStep]:=0;
      end;
      for WorkAreaType:=1 to NumberOfAreaTypes do
        WorkplaceDistribution[AreaType][WorkAreaType][timeStep]:=
        WorkplaceDistribution[AreaType][WorkAreaType][timeStep-1];
  end;

  for AreaType:=1 to NumberOfAreaTypes do
  for WorkerGr:=1 to NumberOfWorkerGrs do
  for IncomeGr:=1 to NumberOfIncomeGrs do
  for EthnicGr:=1 to NumberOfEthnicGrs do
  for HhldType:=1 to NumberOfHhldTypes do
  for AgeGroup:=1 to NumberOfAgeGroups do begin

    residents:=Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1];

    if residents>0 then begin

      SpacePerPerson:=
           BaseResidentialSpacePerPerson[AreaType][HHldType]/(5280.0*5280) {sq feet to sq miles}
        * (Dummy(IncomeGr,1)*LowIncomeEffectOnSpacePerHousehold[timeStep]
          +Dummy(IncomeGr,2)* 1
          +Dummy(IncomeGr,3)*HighIncomeEffectOnSpacePerHousehold[timeStep])
        * ExogenousEffectOnResidentialSpacePerHousehold[timeStep];

      ResidentialSpaceDemand[AreaType][timeStep]:=ResidentialSpaceDemand[AreaType][timeStep]
        + (residents * SpacePerPerson);


      if (WorkerGr=1) {worker} then begin
        autoTrips:=CarDriverWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1];
        transitTrips:=TransitWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1];

        for WorkAreaType:=1 to NumberOfAreaTypes do begin

          commuters := residents * WorkplaceDistribution[AreaType][WorkAreaType][timeStep];

          JobDemand[WorkAreaType][timeStep]:=JobDemand[WorkAreaType][timeStep]
            + commuters;

          {miles by area type - work auto trips}
          for RoadAreaType:=1 to NumberOfAreaTypes do begin

            tdistance:= CarTripAverageODDistance[AreaType][WorkAreaType]
            * ODThroughDistanceFraction[AreaType][WorkAreaType][RoadAreaType];

            if tdistance>0 then
            for RoadType:=1 to NumberOfRoadTypes do begin

              if (RoadType=AreaType) and (RoadType=WorkAreaType) then ODType:=1 else
              if (RoadType=AreaType) or  (RoadType=WorkAreaType) then ODType:=2 else ODType:=3;
              rdistance:=tdistance * DistanceFractionByRoadType[AreaType][ODType][RoadType];

              WorkTripRoadMileDemand[RoadAreaType][RoadType][timeStep]:=
              WorkTripRoadMileDemand[RoadAreaType][RoadType][timeStep]
              + autoTrips * rdistance;
            end;
          end;
          {miles by transit - work transit trips}
          for RoadAreaType:=1 to NumberOfAreaTypes do begin

            tdistance:= TransitTripAverageODDistance[AreaType][WorkAreaType]
            * ODThroughDistanceFraction[AreaType][WorkAreaType][RoadAreaType];

            if tdistance>0 then
            for TransitType:=1 to NumberOfTransitTypes do begin

              if (TransitType=1) then
                rdistance:=tdistance * TransitRailPAFraction[AreaType][WorkAreaType]
              else
                rdistance:=tdistance * (1.0 - TransitRailPAFraction[AreaType][WorkAreaType]);

              WorkTripTransitMileDemand[RoadAreaType][TransitType][timeStep]:=
              WorkTripTransitMileDemand[RoadAreaType][TransitType][timeStep]
              + transitTrips * rdistance;
            end;
          end;
        end;
      end;


      begin {non-work}
        autoTrips:=CarDriverNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1];
        transitTrips:=TransitNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1];

        for DestAreaType:=1 to NumberOfAreaTypes do begin

          {miles by area type - work auto trips}
          for RoadAreaType:=1 to NumberOfAreaTypes do begin

            tdistance:= CarTripAverageODDistance[AreaType][DestAreaType]
            * ODThroughDistanceFraction[AreaType][DestAreaType][RoadAreaType];

            if tdistance>0 then
            for RoadType:=1 to NumberOfRoadTypes do begin

              if (RoadType=AreaType) and (RoadType=DestAreaType) then ODType:=1 else
              if (RoadType=AreaType) or  (RoadType=DestAreaType) then ODType:=2 else ODType:=3;
              rdistance:=tdistance * DistanceFractionByRoadType[AreaType][ODType][RoadType];

              NonWorkTripRoadMileDemand[RoadAreaType][RoadType][timeStep]:=
              NonWorkTripRoadMileDemand[RoadAreaType][RoadType][timeStep]
              + autoTrips * rdistance;
            end;
          end;
          {miles by transit - work transit trips}
          for RoadAreaType:=1 to NumberOfAreaTypes do begin

            tdistance:= TransitTripAverageODDistance[AreaType][DestAreaType]
            * ODThroughDistanceFraction[AreaType][DestAreaType][RoadAreaType];

            if tdistance>0 then
            for TransitType:=1 to NumberOfTransitTypes do begin

              if (TransitType=1) then
                rdistance:=tdistance * TransitRailPAFraction[AreaType][DestAreaType]
              else
                rdistance:=tdistance * (1.0 - TransitRailPAFraction[AreaType][DestAreaType]);

              NonWorkTripTransitMileDemand[RoadAreaType][TransitType][timeStep]:=
              NonWorkTripTransitMileDemand[RoadAreaType][TransitType][timeStep]
              + transitTrips * rdistance;
            end;
          end;
        end;
      end;

    end;
  end;

end;

procedure CalculateEmploymentFeedbacks(timeStep:integer);
var AreaType, EmploymentType: byte;
begin


 {get ratio of jobs to labor force in each area type from the previous time step}
  for AreaType:=1 to NumberOfAreaTypes do begin
    JobSupply[AreaType][timeStep]:=0;
    CommercialSpaceDemand[AreaType][timeStep]:=0;

    for EmploymentType:=1 to NumberOfEmploymentTypes do begin
      JobSupply[AreaType][timeStep]:=JobSupply[AreaType][timeStep]
        +Jobs[AreaType][EmploymentType][timeStep-1];
      CommercialSpaceDemand[AreaType][timeStep]:=CommercialSpaceDemand[AreaType][timeStep]
        +Jobs[AreaType][EmploymentType][timeStep-1]
        *BaseCommercialSpacePerJob[AreaType][EmploymentType]/(5280.0*5280) {sq feet to sq miles}
        *ExogenousEffectOnCommercialSpacePerJob[timeStep];
    end;
  end;

  {do job index relative to period 1, since not all persons who live in area work in area}
  for AreaType:=1 to NumberOfAreaTypes do begin
    JobDemandSupplyIndex[AreaType][timeStep]:=
      (JobDemand[AreaType][timeStep] / Max(1,JobSupply[AreaType][timeStep]))
     /(JobDemand[AreaType][  1     ] / Max(1,JobSupply[AreaType][   1    ]));
  end;

end;

procedure CalculateLandUseFeedbacks(timeStep:integer);
var AreaType: byte;
const LUResidential=2; LUCommercial=1; LUDevelopable=3; LUProtected=4;
begin

 {get ratio of demand and supply for Residential space, commercial space, and developable space}
  for AreaType:=1 to NumberOfAreaTypes do begin
    CommercialSpaceSupply[AreaType][timeStep]:=Land[AreaType][LUCommercial][timeStep-1];
    ResidentialSpaceSupply[AreaType][timeStep]:=Land[AreaType][LUResidential][timeStep-1];
    DevelopableSpaceSupply[AreaType][timeStep]:=Land[AreaType][LUDevelopable][timeStep-1];

    ResidentialSpaceDemandSupplyIndex[AreaType][timeStep]:=
       ResidentialSpaceDemand[AreaType][timeStep]  / Max(1,ResidentialSpaceSupply[AreaType][timeStep]);
    CommercialSpaceDemandSupplyIndex[AreaType][timeStep]:=
       CommercialSpaceDemand[AreaType][timeStep] / Max(1,CommercialSpaceSupply[AreaType][timeStep]);
    DevelopableSpaceDemandSupplyIndex[AreaType][timeStep]:=
       (Max(0,ResidentialSpaceDemand[AreaType][timeStep] - ResidentialSpaceSupply[AreaType][timeStep])
       +Max(0,CommercialSpaceDemand[AreaType][timeStep] - CommercialSpaceSupply[AreaType][timeStep]))
      / Max(1,DevelopableSpaceSupply[AreaType][timeStep] );
  end;

end;

procedure CalculateTransportationSupplyFeedbacks(timeStep:integer);
var AreaType, RoadType, TransitType: byte;
TotalRoadDemand, TotalRoadSupply, TotalTransitDemand, TotalTransitSupply:single;

const RoadTypeWeight:array[1..NumberOfRoadTypes] of single=(0.5,0.4,0.1);
      TransitTypeWeight:array[1..NumberOfTransitTypes] of single=(0.6,0.4);

begin


 {get ratio of demand and supply for road lane miles in each area type and road type}
  for AreaType:=1 to NumberOfAreaTypes do begin

    RoadVehicleCapacityDemandSupplyIndex[AreaType][timeStep]:= 0;

    for RoadType:=1 to NumberOfRoadTypes do begin

      RoadVehicleCapacitySupply[AreaType][RoadType][timeStep]:=
        RoadLaneMiles[AreaType][RoadType][timeStep-1]
      * BaseRoadLaneCapacityPerHour[AreaType,RoadType]
      * ExogenousEffectOnRoadCapacityPerLane[timeStep];


      RoadVehicleCapacityDemand[AreaType][RoadType][timeStep]:=
         WorkTripRoadMileDemand[AreaType][RoadType][timeStep-1]
       * WorkTripPeakHourFraction
       + NonWorkTripRoadMileDemand[AreaType][RoadType][timeStep-1]
       * NonWorkTripPeakHourFraction;


      RoadVehicleCapacityDemandSupplyIndex[AreaType][timeStep]:=
      RoadVehicleCapacityDemandSupplyIndex[AreaType][timeStep]
      + RoadTypeWeight[RoadType]
      * RoadVehicleCapacityDemand[AreaType][RoadType][timeStep]
       /Max(1,RoadVehicleCapacitySupply[AreaType][RoadType][timeStep]);
    end;
  end;

 {get ratio of demand and supply for transit route miles in each area type and transit type}

 for AreaType:=1 to NumberOfAreaTypes do begin

   TransitPassengerCapacityDemandSupplyIndex[AreaType][timeStep]:=0;

   for TransitType:=1 to NumberOfTransitTypes do begin

      TransitPassengerCapacitySupply[AreaType][TransitType][timeStep]:=
      + TransitRouteMiles[AreaType][TransitType][timeStep-1]
      * BaseTransitRouteCapacityPerHour[AreaType,TransitType]
      * ExogenousEffectOnTransitCapacityPerRoute[timeStep];

      TransitPassengerCapacityDemand[AreaType][TransitType][timeStep]:=
         WorkTripTransitMileDemand[AreaType][TransitType][timeStep-1]
       * WorkTripPeakHourFraction
       + NonWorkTripTransitMileDemand[AreaType][TransitType][timeStep-1]
       * NonWorkTripPeakHourFraction;

     TransitPassengerCapacityDemandSupplyIndex[AreaType][timeStep]:=
     TransitPassengerCapacityDemandSupplyIndex[AreaType][timeStep]+
     + TransitTypeWeight[TransitType]
     * TransitPassengerCapacityDemand[AreaType][TransitType][timeStep]
      /Max(1,TransitPassengerCapacitySupply[AreaType][TransitType][timeStep]);
   end;
 end;
end;


procedure CalculateEmploymentTransitionRates(timeStep:integer);
var AreaType, EmploymentType, AreaType2: byte;
    EmployerAttractivenessIndex, ExternalEmployerAttractivenessIndex
    :array[1..NumberOfAreaTypes,1..NumberOfEmploymentTypes] of single;
    CurrentJobs,JobsMoved:single;
begin

  {set attractivness index for employment}
  for AreaType:=1 to NumberOfAreaTypes do
  for EmploymentType:=1 to NumberOfEmploymentTypes do begin

    EmployerAttractivenessIndex[AreaType][EmploymentType]:=

     ( EffectCurve(C_EffectOfJobDemandSupplyIndexOnEmployerAttractiveness,
         JobDemandSupplyIndex[AreaType][timeStep])
      * WeightOfJobDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]

     +EffectCurve(C_EffectOfCommercialSpaceDemandSupplyIndexOnEmployerAttractiveness,
         CommercialSpaceDemandSupplyIndex[AreaType][timeStep])
      * WeightOfCommercialSpaceDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]

     +EffectCurve(C_EffectOfRoadMileDemandSupplyIndexOnEmployerAttractiveness,
         RoadVehicleCapacityDemandSupplyIndex[AreaType][timeStep])
      * WeightOfRoadMileDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]
     )/
      ( WeightOfJobDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]
      + WeightOfCommercialSpaceDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]
      + WeightOfRoadMileDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]);

    ExternalEmployerAttractivenessIndex[AreaType][EmploymentType]:=

     ( EffectCurve(C_EffectOfJobDemandSupplyIndexOnEmployerAttractiveness,
         ExternalJobDemandSupplyIndex[timeStep])
      * WeightOfJobDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]

     +EffectCurve(C_EffectOfCommercialSpaceDemandSupplyIndexOnEmployerAttractiveness,
         ExternalCommercialSpaceDemandSupplyIndex[timeStep])
      * WeightOfCommercialSpaceDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]

     +EffectCurve(C_EffectOfRoadMileDemandSupplyIndexOnEmployerAttractiveness,
         ExternalRoadMileDemandSupplyIndex[timeStep])
      * WeightOfRoadMileDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]
     )/
      ( WeightOfJobDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]
      + WeightOfCommercialSpaceDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]
      + WeightOfRoadMileDemandSupplyIndexInEmployerAttractiveness[AreaType][EmploymentType]);


  end;


  for AreaType:=1 to NumberOfAreaTypes do
  for EmploymentType:=1 to NumberOfEmploymentTypes do begin
    JobsCreated[AreaType][EmploymentType][timeStep]:=0;
    JobsLost[AreaType][EmploymentType][timeStep]:=0;
    JobsMovedOut[AreaType][EmploymentType][timeStep]:=0;
    JobsMovedIn[AreaType][EmploymentType][timeStep]:=0;
  end;

  {loop on cells and set rates}
  for AreaType:=1 to NumberOfAreaTypes do
  for EmploymentType:=1 to NumberOfEmploymentTypes do begin


    CurrentJobs:=Jobs[AreaType][EmploymentType][timeStep-1];

    JobsMoved:=CurrentJobs
        * (EmployerAttractivenessIndex[AreaType][EmploymentType]
          -ExternalEmployerAttractivenessIndex[AreaType][EmploymentType]);

    if JobsMoved>0 then begin
         JobsCreated[AreaType][EmploymentType][timeStep]:=
          Min(JobsMoved,CurrentJobs) * TimeStepLength/JobCreationDelay[timeStep]
        * ExogenousEffectOnJobCreationRate[timeStep];
    end
    else begin
         JobsLost[AreaType][EmploymentType][timeStep]:=
           Min(-JobsMoved,CurrentJobs) * TimeStepLength/JobLossDelay[timeStep]
        * ExogenousEffectOnJobLossRate[timeStep];
    end;

    {check other area types, and move jobs there if more attractive}
    for AreaType2:=1 to NumberOfAreaTypes do
    if (AreaType2 <> AreaType) then begin

        JobsMoved:=CurrentJobs
        * (EmployerAttractivenessIndex[AreaType2][EmploymentType]
          -EmployerAttractivenessIndex[AreaType][EmploymentType])
        * ExogenousEffectOnJobMoveRate[timeStep];

        if JobsMoved>0 then begin

          JobsMovedOut[AreaType][EmploymentType][timeStep]:=
            JobsMovedOut[AreaType][EmploymentType][timeStep]
            + Min(JobsMoved,CurrentJobs) * TimeStepLength/JobMoveDelay[timeStep];

          JobsMovedIn[AreaType2][EmploymentType][timeStep]:=
            JobsMovedIn[AreaType2][EmploymentType][timeStep]
            + Min(JobsMoved,CurrentJobs) * TimeStepLength/JobMoveDelay[timeStep];
       end;
    end;
  end;

end; {CalculateEmploymentTransitionRates}


procedure CalculateLandUseTransitionRates(timeStep:integer);
var AreaType : byte;
    NewResidentialSpaceNeeded,ExcessResidentialSpace,NewResidentialSpaceDeveloped,ResidentialSpaceReleased,
    NewCommercialSpaceNeeded,ExcessCommercialSpace,NewCommercialSpaceDeveloped,CommercialSpaceReleased,
    ProtectedSpaceReleased,DevelopableResidentialSpace,DevelopableCommercialSpace,
    IndicatedResidentialDevelopment,IndicatedCommercialDevelopment,DevelopableLandSufficiencyFraction:single;

const LUResidential=2; LUCommercial=1; LUDevelopable=3; LUProtected=4;
begin

 for AreaType:=1 to NumberOfAreaTypes do begin

  NewResidentialSpaceNeeded:= Max(0,ResidentialSpaceDemand[AreaType][timeStep] - ResidentialSpaceSupply[AreaType][timeStep]);

  ExcessResidentialSpace:=Max(0,ResidentialSpaceSupply[AreaType][timeStep] - ResidentialSpaceDemand[AreaType][timeStep]);

  DevelopableResidentialSpace:=DevelopableSpaceSupply[AreaType][TimeStep]
    * FractionOfDevelopableLandAllowedForResidential[AreaType];

  IndicatedResidentialDevelopment:=Min(NewResidentialSpaceNeeded,DevelopableResidentialSpace);

  NewCommercialSpaceNeeded:= Max(0,CommercialSpaceDemand[AreaType][timeStep] - CommercialSpaceSupply[AreaType][timeStep]);

  ExcessCommercialSpace:=Max(0,CommercialSpaceSupply[AreaType][timeStep] - CommercialSpaceDemand[AreaType][timeStep]);

  DevelopableCommercialSpace:=DevelopableSpaceSupply[AreaType][TimeStep]
     * FractionOfDevelopableLandAllowedForCommercial[AreaType];

  IndicatedCommercialDevelopment:=Min(NewCommercialSpaceNeeded,DevelopableCommercialSpace);

  DevelopableLandSufficiencyFraction:= DevelopableSpaceSupply[AreaType][TimeStep]/
     Max(1.0,IndicatedResidentialDevelopment+IndicatedCommercialDevelopment);

  if DevelopableLandSufficiencyFraction<1.0 then begin
      IndicatedResidentialDevelopment:=IndicatedResidentialDevelopment
        * DevelopableLandSufficiencyFraction;
      IndicatedCommercialDevelopment:=IndicatedCommercialDevelopment
        * DevelopableLandSufficiencyFraction;
  end;

  if NewResidentialSpaceNeeded>0 then
    NewResidentialSpaceDeveloped:=IndicatedResidentialDevelopment
     * TimeStepLength/ResidentialSpaceDevelopmentDelay[timeStep]
  else NewResidentialSpaceDeveloped:=0;

  if ExcessResidentialSpace>0 then
    ResidentialSpaceReleased:=ExcessResidentialSpace
     * TimeStepLength/ResidentialSpaceReleaseDelay[timeStep]
  else ResidentialSpaceReleased:=0;

  if NewCommercialSpaceNeeded>0 then
    NewCommercialSpaceDeveloped:=IndicatedCommercialDevelopment
     * TimeStepLength/CommercialSpaceDevelopmentDelay[timeStep]
   else NewCommercialSpaceDeveloped:=0;

  if ExcessCommercialSpace>0 then
    CommercialSpaceReleased:=ExcessCommercialSpace
     * TimeStepLength/CommercialSpaceReleaseDelay[timeStep]
  else CommercialSpaceReleased:=0;

  ProtectedSpaceReleased:=  {this can be negative - added to protection}
    Land[AreaType][LUProtected][timeStep-1] *
    (1.0 - ExogenousEffectOnLandProtection[timeStep])
     * TimeStepLength / LandProtectionProcessDelay[timeStep];

  ChangeInLandUseIn[AreaType][LUResidential][timeStep]:=NewResidentialSpaceDeveloped;
  ChangeInLandUseOut[AreaType][LUResidential][timeStep]:=ResidentialSpaceReleased;

  ChangeInLandUseIn[AreaType][LUCommercial][timeStep]:=NewCommercialSpaceDeveloped;
  ChangeInLandUseOut[AreaType][LUCommercial][timeStep]:=CommercialSpaceReleased;

  ChangeInLandUseIn[AreaType][LUDevelopable][timeStep]:=ResidentialSpaceReleased + CommercialSpaceReleased + ProtectedSpaceReleased;;
  ChangeInLandUseOut[AreaType][LUDevelopable][timeStep]:=NewResidentialSpaceDeveloped + NewCommercialSpaceDeveloped;

  ChangeInLandUseOut[AreaType][LUProtected][timeStep]:=ProtectedSpaceReleased;

 end;

end; {CalculateLandUseTransitionRates}

procedure CalculateTransportationSupplyTransitionRates(timeStep:integer);
var AreaType, RoadType, TransitType : byte;
FractionNewRoadMilesNeeded, FractionNewRoadMilesAdded, FractionNewTransitMilesNeeded, FractionNewTransitMilesAdded:single;
begin

 for AreaType:=1 to NumberOfAreaTypes do begin

   for RoadType:=1 to NumberOfRoadTypes do begin

    FractionNewRoadMilesNeeded:= Max(0,RoadVehicleCapacityDemand[AreaType][RoadType][timeStep] /
    Max(1,RoadVehicleCapacitySupply[AreaType][RoadType][timeStep]) - 1.0);

    if FractionNewRoadMilesNeeded>0 then begin
      FractionNewRoadMilesAdded:=FractionNewRoadMilesNeeded
      * ExogenousEffectOnRoadCapacityAddition[timeStep]
      * TimeStepLength/RoadCapacityAdditionDelay[timeStep];

      RoadLaneMilesAdded[AreaType][RoadType][timeStep]:=
          RoadLaneMiles[AreaType][RoadType][timeStep-1]
         * FractionNewRoadMilesAdded;
    end else begin

      RoadLaneMilesLost[AreaType][RoadType][timeStep]:=
        RoadLaneMiles[AreaType][RoadType][timeStep-1]
        * TimeStepLength/RoadCapacityRetirementDelay[timeStep];
    end;
   end;

   for TransitType:=1 to NumberOfTransitTypes do begin

     FractionNewTransitMilesNeeded:= Max(0,TransitPassengerCapacityDemand[AreaType][TransitType][timeStep] /
     Max(1,TransitPassengerCapacitySupply[AreaType][TransitType][timeStep]) -1.0);

     if FractionNewTransitMilesNeeded>0 then begin
        FractionNewTransitMilesAdded:=FractionNewTransitMilesNeeded
        * ExogenousEffectOnTransitCapacityAddition[timeStep]
        * TimeStepLength/TransitCapacityAdditionDelay[timeStep];

        TransitRouteMilesAdded[AreaType][TransitType][timeStep]:=
          TransitRouteMiles[AreaType][TransitType][timeStep-1]
          * FractionNewTransitMilesAdded;
      end else begin

       TransitRouteMilesLost[AreaType][TransitType][timeStep]:=
          TransitRouteMiles[AreaType][TransitType][timeStep-1]
          * TimeStepLength/TransitCapacityRetirementDelay[timeStep];
      end;
    end;
  end;

end; {CalculateTransportationSuppplyTransitionRates}


procedure CalculateDemographicTransitionRates(timeStep:integer);
var PreviousPopulation, ResidentsMoved,NewHHChildrenFraction,
tempSingle,tempCouple,temp1,temp2, PeopleMoved:single;
{indices}
AreaType,AreaType2,
WorkerGr,
IncomeGr,
EthnicGr,
HhldType,
AgeGroup,
MigrationType,
NewAreaType,
NewWorkerGr,
NewIncomeGr,
NewEthnicGr,
NewHhldType,
NewAgeGroup,
BirthEthnicGr : byte;
ResidentAttractivenessIndex,ExternalResidentAttractivenessIndex
:array[1..NumberOfAreaTypes,1..NumberOfMigrationTypes] of single;

begin
  {set attractivness index for residents}
  for AreaType:=1 to NumberOfAreaTypes do
  for MigrationType:=1 to NumberOfMigrationTypes do begin

    ResidentAttractivenessIndex[AreaType][MigrationType]:=

     ( EffectCurve(C_EffectOfJobDemandSupplyIndexOnResidentAttractiveness,
         JobDemandSupplyIndex[AreaType][timeStep])
      * WeightOfJobDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]

     +EffectCurve(C_EffectOfResidentialSpaceDemandSupplyIndexOnResidentAttractiveness,
         ResidentialSpaceDemandSupplyIndex[AreaType][timeStep])
      * WeightOfResidentialSpaceDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]

     +EffectCurve(C_EffectOfRoadMileDemandSupplyIndexOnResidentAttractiveness,
         RoadVehicleCapacityDemandSupplyIndex[AreaType][timeStep])
      * WeightOfRoadMileDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]
     )/
      ( WeightOfJobDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]
      + WeightOfResidentialSpaceDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]
      + WeightOfRoadMileDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]);

    ExternalResidentAttractivenessIndex[AreaType][MigrationType]:=

     ( EffectCurve(C_EffectOfJobDemandSupplyIndexOnResidentAttractiveness,
         ExternalJobDemandSupplyIndex[timeStep])
      * WeightOfJobDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]

     +EffectCurve(C_EffectOfResidentialSpaceDemandSupplyIndexOnResidentAttractiveness,
         ExternalResidentialSpaceDemandSupplyIndex[timeStep])
      * WeightOfResidentialSpaceDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]

     +EffectCurve(C_EffectOfRoadMileDemandSupplyIndexOnResidentAttractiveness,
         ExternalRoadMileDemandSupplyIndex[timeStep])
      * WeightOfRoadMileDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]
     )/
      ( WeightOfJobDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]
      + WeightOfResidentialSpaceDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]
      + WeightOfRoadMileDemandSupplyIndexInResidentAttractiveness[AreaType][MigrationType]);
  end;

  {initialize all entries for each cell to 0}
  for AreaType:=1 to NumberOfAreaTypes do
  for WorkerGr:=1 to NumberOfWorkerGrs do
  for IncomeGr:=1 to NumberOfIncomeGrs do
  for EthnicGr:=1 to NumberOfEthnicGrs do
  for HhldType:=1 to NumberOfHhldTypes do
  for AgeGroup:=1 to NumberOfAgeGroups do begin
        BirthsFrom[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        DeathsOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        MarriagesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        DivorcesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        FirstChildOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        EmptyNestOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        LeaveNestOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        AcculturationOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        WorkerStatusOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        IncomeGroupOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        BirthsIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        MarriagesIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        DivorcesIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        FirstChildIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        EmptyNestIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        LeaveNestIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        AcculturationIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        WorkerStatusIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        IncomeGroupIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        ForeignInmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        DomesticInmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        RegionalInmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        ForeignOutmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        DomesticOutmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        RegionalOutmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
  end;

  {apply rates for each cell}
  for AreaType:=1 to NumberOfAreaTypes do
  for WorkerGr:=1 to NumberOfWorkerGrs do
  for IncomeGr:=1 to NumberOfIncomeGrs do
  for EthnicGr:=1 to NumberOfEthnicGrs do
  for HhldType:=1 to NumberOfHhldTypes do
  for AgeGroup:=1 to NumberOfAgeGroups do begin

     PreviousPopulation:=Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1];

     if PreviousPopulation > 0 then begin

         {Calculate number ageing to the next age group}
        if (AgeGroupDuration[AgeGroup]>0.5) then begin
          {ageing rate is based only on duration of age cohort}
          AgeingOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=PreviousPopulation
          * TimeStepLength / AgeGroupDuration[AgeGroup];
         {put them into next age group}
          NewAgeGroup:=AgeGroup+1;
          AgeingIn[AreaType][NewAgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=AgeingIn[AreaType][NewAgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          + AgeingOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep];
        end;

       {Calculate number of deaths}
        DeathsOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
        :=PreviousPopulation
        * BaseMortalityRate[AgeGroup][HhldType][EthnicGr] * TimeStepLength
        * (LowIncomeDummy[IncomeGr] * LowIncomeEffectOnMortalityRate[timeStep]
         + MiddleIncomeDummy[IncomeGr]
         + HighIncomeDummy[IncomeGr] * HighIncomeEffectOnMortalityRate[timeStep])
        * ExogenousEffectOnMortalityRate[timeStep];
         {deaths aren't put into any other group}

       {Calculate number of births}
        BirthsFrom[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
        :=PreviousPopulation
        * BaseFertilityRate[AgeGroup][HhldType][EthnicGr] * TimeStepLength
        * (LowIncomeDummy[IncomeGr] * LowIncomeEffectOnFertilityRate[timeStep]
         + MiddleIncomeDummy[IncomeGr]
         + HighIncomeDummy[IncomeGr] * HighIncomeEffectOnFertilityRate[timeStep])
        * ExogenousEffectOnFertilityRate[timeStep];

        {If first child, all adults in HH become "full nest" household}
        if (NumberOfChildren[HhldType]<0.5) then begin
          FirstChildOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          := BirthsFrom[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] * NumberOfAdults[HhldType];

          NewHhldType:=HhldType + 2; {same number of adults, 1+ kids}

         {add full nest to new hhld type}
          FirstChildIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=FirstChildIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          + BirthsFrom[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] * NumberOfAdults[HhldType];

        end else begin
          NewHhldType:=HhldType; {not first child, same hhld type}
        end;

        BirthEthnicGr:=BirthEthnicGroup[EthnicGr];
        BirthsIn[AreaType][BirthAgeGroup][NewHhldType][BirthEthnicGr][IncomeGr][BirthWorkerGr][timeStep]
        :=BirthsIn[AreaType][BirthAgeGroup][NewHhldType][BirthEthnicGr][IncomeGr][BirthWorkerGr][timeStep]
          + BirthsFrom[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep];

        {Calculate number of "marriages"}
        if (NumberOfAdults[HhldType]<1.99) then begin
          MarriagesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=PreviousPopulation
          * BaseMarriageRate[AgeGroup][HhldType][EthnicGr] * TimeStepLength
          * (LowIncomeDummy[IncomeGr] * LowIncomeEffectOnMarriageRate[timeStep]
          +  MiddleIncomeDummy[IncomeGr]
          +  HighIncomeDummy[IncomeGr] * HighIncomeEffectOnMarriageRate[timeStep])
          * ExogenousEffectOnMarriageRate[timeStep];

          if NumberOfChildren[HhldType]=0
            then NewHHChildrenFraction:=MarryNoChildren_ChildrenFraction
            else NewHHChildrenFraction:=MarryHasChildren_ChildrenFraction;

         {add marriages to new hhld types}
          NewHhldType:=3; {couple, no children}
          MarriagesIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=MarriagesIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          + MarriagesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          *(1.0-NewHHChildrenFraction);

          NewHhldType:=4; {couple, children}
          MarriagesIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=MarriagesIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          + MarriagesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          * NewHHChildrenFraction;
        end;

       {Calculate number "divorces"}
        if (NumberOfAdults[HhldType]>1.99) then begin
          DivorcesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=PreviousPopulation
          * BaseDivorceRate[AgeGroup][HhldType][EthnicGr] * TimeStepLength
          * (LowIncomeDummy[IncomeGr] * LowIncomeEffectOnDivorceRate[timeStep]
          +  MiddleIncomeDummy[IncomeGr]
          +  HighIncomeDummy[IncomeGr] * HighIncomeEffectOnDivorceRate[timeStep])
          * ExogenousEffectOnDivorceRate[timeStep];

          if NumberOfChildren[HhldType]=0
            then NewHHChildrenFraction:=DivorceNoChildren_ChildrenFraction
            else NewHHChildrenFraction:=DivorceHasChildren_ChildrenFraction;

         {add divorces to new hhld types}
          NewHhldType:=1; {single, no children}
          DivorcesIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=DivorcesIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          + DivorcesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          *(1.0-NewHHChildrenFraction);

         {add divorces to new hhld types}
          NewHhldType:=2; {single, w/ children}
          DivorcesIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=DivorcesIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          + DivorcesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          * NewHHChildrenFraction;
        end;

        {Calculate number of 1+ child HH transitioning to 0 child ("empty nest" }
        if (NumberOfChildren[HhldType]>0.5) then begin
          EmptyNestOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=PreviousPopulation
          * BaseEmptyNestRate[AgeGroup][HhldType][EthnicGr] * TimeStepLength
          * (LowIncomeDummy[IncomeGr] * LowIncomeEffectOnEmptyNestRate[timeStep]
          +  MiddleIncomeDummy[IncomeGr]
          +  HighIncomeDummy[IncomeGr] * HighIncomeEffectOnEmptyNestRate[timeStep])
          * ExogenousEffectOnEmptyNestRate[timeStep];

         {add to new hhld type}
          NewHhldType:=HhldType-2; {same adults, no children}
           EmptyNestIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=EmptyNestIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          + EmptyNestOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep];
        end;

        {calculate number of children "leaving the nest" }
        if (NumberOfChildren[HhldType]>0.5) then begin
          tempSingle:=PreviousPopulation
          * BaseLeaveNestSingleRate[AgeGroup][HhldType][EthnicGr] * TimeStepLength
          * (LowIncomeDummy[IncomeGr] * LowIncomeEffectOnEmptyNestRate[timeStep]
          +  MiddleIncomeDummy[IncomeGr]
          +  HighIncomeDummy[IncomeGr] * HighIncomeEffectOnEmptyNestRate[timeStep])
          * ExogenousEffectOnEmptyNestRate[timeStep];
          tempCouple:=PreviousPopulation
          * BaseLeaveNestCoupleRate[AgeGroup][HhldType][EthnicGr] * TimeStepLength
          * (LowIncomeDummy[IncomeGr] * LowIncomeEffectOnEmptyNestRate[timeStep]
          +  MiddleIncomeDummy[IncomeGr]
          +  HighIncomeDummy[IncomeGr] * HighIncomeEffectOnEmptyNestRate[timeStep])
          * ExogenousEffectOnEmptyNestRate[timeStep];

          LeaveNestOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
            := tempSingle + tempCouple;

        {add to new hhld types}
          NewHhldType:=1; {single, no children}
          LeaveNestIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=LeaveNestIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
           + tempSingle * (1.0-LeaveNestSingle_ChildrenFraction);

          NewHhldType:=2; {couple, no children}
          LeaveNestIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=LeaveNestIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
           + tempCouple * (1.0-LeaveNestCouple_ChildrenFraction);

          NewHhldType:=3; {single, w/ children}
          LeaveNestIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=LeaveNestIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
           + tempSingle * LeaveNestSingle_ChildrenFraction;

          NewHhldType:=4; {couple, w/ children}
          LeaveNestIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=LeaveNestIn[AreaType][AgeGroup][NewHhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
           + tempCouple * LeaveNestCouple_ChildrenFraction;
        end;


      {Calculate workforce shifts}
        if WorkerGr=1 then begin {in workforce}
          WorkerStatusOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=PreviousPopulation
          * BaseLeaveWorkforceRate[AgeGroup][HhldType][EthnicGr]
          * TimeStepLength / WorkforceChangeDelay[timeStep]
          * ExogenousEffectOnLeaveWorkforceRate[timeStep];

          NewWorkerGr:=2; {out of workforce}
          WorkerStatusIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][NewWorkerGr][timeStep]
          :=WorkerStatusIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][NewWorkerGr][timeStep]
          + WorkerStatusOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep];
        end;
        if WorkerGr=2 then begin {out workforce}
          WorkerStatusOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=PreviousPopulation
          * BaseEnterWorkforceRate[AgeGroup][HhldType][EthnicGr]
          * TimeStepLength / WorkforceChangeDelay[timeStep]
          * ExogenousEffectOnEnterWorkforceRate[timeStep];

          NewWorkerGr:=1; {in workforce}
          WorkerStatusIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][NewWorkerGr][timeStep]
          :=WorkerStatusIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][NewWorkerGr][timeStep]
          + WorkerStatusOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep];
        end;

      {Calculate income shifts}
        if IncomeGr=1 then begin {leave low income}
          IncomeGroupOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=PreviousPopulation
          * BaseLeaveLowIncomeRate[AgeGroup][HhldType][EthnicGr]
          * TimeStepLength / IncomeChangeDelay[timeStep]
          * ExogenousEffectOnLeaveLowIncomeRate[timeStep];

          NewIncomeGr:=2; {enter middle income}
          IncomeGroupIn[AreaType][AgeGroup][HhldType][EthnicGr][NewIncomeGr][WorkerGr][timeStep]
          :=IncomeGroupIn[AreaType][AgeGroup][HhldType][EthnicGr][NewIncomeGr][WorkerGr][timeStep]
          + IncomeGroupOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep];
        end else
        if IncomeGr=2 then begin {leave middle income to low}
          temp1
          :=PreviousPopulation
          * BaseEnterLowIncomeRate[AgeGroup][HhldType][EthnicGr]
          * TimeStepLength / IncomeChangeDelay[timeStep]
          * ExogenousEffectOnEnterLowIncomeRate[timeStep];

          NewIncomeGr:=1; {enter low income}
          IncomeGroupIn[AreaType][AgeGroup][HhldType][EthnicGr][NewIncomeGr][WorkerGr][timeStep]
          :=IncomeGroupIn[AreaType][AgeGroup][HhldType][EthnicGr][NewIncomeGr][WorkerGr][timeStep]
          + temp1;

          {leave middle income to high}
          temp2
          :=PreviousPopulation
          * BaseEnterHighIncomeRate[AgeGroup][HhldType][EthnicGr]
          * TimeStepLength / IncomeChangeDelay[timeStep]
          * ExogenousEffectOnEnterHighIncomeRate[timeStep];

          NewIncomeGr:=3; {enter high income}
          IncomeGroupIn[AreaType][AgeGroup][HhldType][EthnicGr][NewIncomeGr][WorkerGr][timeStep]
          :=IncomeGroupIn[AreaType][AgeGroup][HhldType][EthnicGr][NewIncomeGr][WorkerGr][timeStep]
          + temp2;

          IncomeGroupOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=temp1+temp2;
         end else
         if IncomeGr=3 then begin {leave high income}
          IncomeGroupOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=PreviousPopulation
          * BaseLeaveHighIncomeRate[AgeGroup][HhldType][EthnicGr]
          * TimeStepLength / IncomeChangeDelay[timeStep]
          * ExogenousEffectOnLeaveHighIncomeRate[timeStep];

          NewIncomeGr:=2; {enter middle income}
          IncomeGroupIn[AreaType][AgeGroup][HhldType][EthnicGr][NewIncomeGr][WorkerGr][timeStep]
          :=IncomeGroupIn[AreaType][AgeGroup][HhldType][EthnicGr][NewIncomeGr][WorkerGr][timeStep]
          + IncomeGroupOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep];
        end;

      {Calculate number of non-US Born reaching 20 years in US ("acculturation") }
        if (EthnicGrDuration[EthnicGr]<0.1) then begin
          AcculturationOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=0;
        end else begin
          AcculturationOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
          :=PreviousPopulation
          * TimeStepLength / EthnicGrDuration[EthnicGr];

          NewEthnicGr:=NextEthnicGroup[EthnicGr];
         {add acculturated to new ethnic gr}
          AcculturationIn[AreaType][AgeGroup][HhldType][NewEthnicGr][IncomeGr][WorkerGr][timeStep]
          :=AcculturationIn[AreaType][AgeGroup][HhldType][NewEthnicGr][IncomeGr][WorkerGr][timeStep]
          + AcculturationOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep];
        end;

        {Foreign migration only in foreign born <20 years ethnic group}
        if (EthnicGrDuration[EthnicGr]>0.1) then begin

          MigrationType:=1;

          PeopleMoved:=PreviousPopulation
          * BaseForeignInmigrationRate * MigrationRateMultiplier[AgeGroup][HhldType][EthnicGr]
          * ResidentAttractivenessIndex[AreaType][MigrationType]
          * ExogenousEffectOnForeignInmigrationRate[timeStep];

          ForeignInmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
            PeopleMoved * TimeStepLength/ForeignInmigrationDelay[timeStep];

          PeopleMoved:=PreviousPopulation
          * BaseForeignOutmigrationRate * MigrationRateMultiplier[AgeGroup][HhldType][EthnicGr]
          * (1.0/ResidentAttractivenessIndex[AreaType][MigrationType])
          * ExogenousEffectOnForeignOutmigrationRate[timeStep];

          ForeignOutmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
             PeopleMoved * TimeStepLength/ForeignOutmigrationDelay[timeStep];
        end;

        {Domestic migration}
        begin

          MigrationType:=2;

          PeopleMoved:=PreviousPopulation
          * BaseDomesticMigrationRate  * MigrationRateMultiplier[AgeGroup][HhldType][EthnicGr]
          *(ResidentAttractivenessIndex[AreaType][MigrationType]
          / ExternalResidentAttractivenessIndex[AreaType][MigrationType])
          * ExogenousEffectOnDomesticMigrationRate[timeStep];


          DomesticInmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
             PeopleMoved * TimeStepLength/DomesticMigrationDelay[timeStep];

          PeopleMoved:=PreviousPopulation
          * BaseDomesticMigrationRate  * MigrationRateMultiplier[AgeGroup][HhldType][EthnicGr]
          *(ExternalResidentAttractivenessIndex[AreaType][MigrationType]
          / ResidentAttractivenessIndex[AreaType][MigrationType])
          * ExogenousEffectOnDomesticMigrationRate[timeStep];


          DomesticOutmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
             PeopleMoved * TimeStepLength/DomesticMigrationDelay[timeStep];

        end;

        {Internal regonal migration between area types}

        MigrationType:=3;

        {check other area types, and move jobs there if more attractive}
        for AreaType2:=1 to NumberOfAreaTypes do
        if (AreaType2 <> AreaType) then begin

          PeopleMoved:=PreviousPopulation
          * BaseRegionalMigrationRate
          *(ResidentAttractivenessIndex[AreaType2][MigrationType]
          / ResidentAttractivenessIndex[AreaType][MigrationType])
          * ExogenousEffectOnRegionalMigrationRate[timeStep];

          RegionalOutmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
          RegionalOutmigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
            + PeopleMoved * TimeStepLength/RegionalMigrationDelay[timeStep];

          RegionalInmigration[AreaType2][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
          RegionalInmigration[AreaType2][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
            + PeopleMoved * TimeStepLength/RegionalMigrationDelay[timeStep];
         end;

     end; {population > 0}
  end; {loop on cells}

end; {CalculateDemographicTransitionRates}


procedure ApplyDemographicTransitionRates(timeStep:integer);
var
{indices}
FAreaType,
AreaType,
WorkerGr,
IncomeGr,
EthnicGr,
HhldType,
AgeGroup: byte;
PLeave,PEnter,UFactor,HUrbEnter,HUrbLeave,IUrbEnter,IUrbLeave,PGrowth, ExogPopChange:single;

begin
      {apply transition rates for each cell}
      for AreaType:=1 to NumberOfAreaTypes do
      for WorkerGr:=1 to NumberOfWorkerGrs do
      for IncomeGr:=1 to NumberOfIncomeGrs do
      for EthnicGr:=1 to NumberOfEthnicGrs do
      for HhldType:=1 to NumberOfHhldTypes do
      for AgeGroup:=1 to NumberOfAgeGroups do begin

       {adjust population growth rate first for exogenous change}
       if AreaType=1 then PGrowth:=ExogenousPopulationChangeRate1[timeStep]-ExogenousPopulationChangeRate1[timeStep-1] else
       if AreaType=2 then PGrowth:=ExogenousPopulationChangeRate2[timeStep]-ExogenousPopulationChangeRate2[timeStep-1] else
       if AreaType=3 then PGrowth:=ExogenousPopulationChangeRate3[timeStep]-ExogenousPopulationChangeRate3[timeStep-1] else
       if AreaType=4 then PGrowth:=ExogenousPopulationChangeRate4[timeStep]-ExogenousPopulationChangeRate4[timeStep-1] else
       if AreaType=5 then PGrowth:=ExogenousPopulationChangeRate5[timeStep]-ExogenousPopulationChangeRate5[timeStep-1] else
       if AreaType=6 then PGrowth:=ExogenousPopulationChangeRate6[timeStep]-ExogenousPopulationChangeRate6[timeStep-1] else
       if AreaType=7 then PGrowth:=ExogenousPopulationChangeRate7[timeStep]-ExogenousPopulationChangeRate7[timeStep-1] else
       if AreaType=8 then PGrowth:=ExogenousPopulationChangeRate8[timeStep]-ExogenousPopulationChangeRate8[timeStep-1] else
       if AreaType=9 then PGrowth:=ExogenousPopulationChangeRate9[timeStep]-ExogenousPopulationChangeRate9[timeStep-1] else
       if AreaType=10 then PGrowth:=ExogenousPopulationChangeRate10[timeStep]-ExogenousPopulationChangeRate10[timeStep-1] else
       if AreaType=11 then PGrowth:=ExogenousPopulationChangeRate11[timeStep]-ExogenousPopulationChangeRate11[timeStep-1] else
       if AreaType=12 then PGrowth:=ExogenousPopulationChangeRate12[timeStep]-ExogenousPopulationChangeRate12[timeStep-1];

       ExogPopChange:=PGrowth*Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][0];

       FAreaType:=4*(AreaTypeSubregion[AreaType]-1)+1;

       {apply household type factors toward urban areas}
       HUrbLeave:=0;
       HUrbEnter:=0;
       if HhldType=1 then UFactor:=SingleNoKidsEffectOnMoveTowardsUrbanAreas[timeStep] else
       if HhldType=2 then UFactor:=CoupleNoKidsEffectOnMoveTowardsUrbanAreas[timeStep] else
       if HhldType=3 then UFactor:=SingleWiKidsEffectOnMoveTowardsUrbanAreas[timeStep] else
       if HhldType=4 then UFactor:=CoupleWiKidsEffectOnMoveTowardsUrbanAreas[timeStep];
       if UFactor>1.0 then begin
         PLeave:=(UFactor-1.0)*
         (Population[FAreaType+0][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]
         +Population[FAreaType+1][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]);
         PEnter:=
         (Population[FAreaType+2][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]
         +Population[FAreaType+3][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]);
         if (PLeave>0) and (PEnter>0) then begin
           if (AreaType=FAreaType+0) or (AreaType=FAreaType+1) then
             HUrbLeave:=(UFactor-1.0)*
             Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1] else
           if (AreaType=FAreaType+2) or (AreaType=FAreaType+3) then
             HUrbEnter:=PLeave*
             Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]/PEnter;
         end;
       end else
       if UFactor<1.0 then begin
         PLeave:=(1.0-UFactor)*
         (Population[FAreaType+2][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]
         +Population[FAreaType+3][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]);
         PEnter:=
         (Population[FAreaType+0][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]
         +Population[FAreaType+1][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]);
         if (PLeave>0) and (PEnter>0) then begin
           if (AreaType=FAreaType+2) or (AreaType=FAreaType+3) then
             HUrbLeave:=(1.0-UFactor)*
             Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1] else
           if (AreaType=FAreaType+0) or (AreaType=FAreaType+1) then
             HUrbEnter:=PLeave*
             Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]/PEnter;
         end;
       end;

       {apply income factors toward urban areas}
       IUrbLeave:=0;
       IUrbEnter:=0;
       if IncomeGr=1 then UFactor:=LowIncomeEffectOnMoveTowardsUrbanAreas[timeStep] else
       if IncomeGr=2 then UFactor:=1.0 else
       if IncomeGr=3 then UFactor:=HighIncomeEffectOnMoveTowardsUrbanAreas[timeStep];
       if UFactor>1.0 then begin
         PLeave:=(UFactor-1.0)*
         (Population[FAreaType+0][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]
         +Population[FAreaType+1][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]);
         PEnter:=
         (Population[FAreaType+2][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]
         +Population[FAreaType+3][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]);
         if (PLeave>0) and (PEnter>0) then begin
           if (AreaType=FAreaType+0) or (AreaType=FAreaType+1) then
             IUrbLeave:=(UFactor-1.0)*
             Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1] else
           if (AreaType=FAreaType+2) or (AreaType=FAreaType+3) then
             IUrbEnter:=PLeave*
             Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]/PEnter;
         end;
       end else
       if UFactor<1.0 then begin
         PLeave:=(1.0-UFactor)*
         (Population[FAreaType+2][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]
         +Population[FAreaType+3][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]);
         PEnter:=
         (Population[FAreaType+0][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]
         +Population[FAreaType+1][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]);
         if (PLeave>0) and (PEnter>0) then begin
           if (AreaType=FAreaType+2) or (AreaType=FAreaType+3) then
             IUrbLeave:=(1.0-UFactor)*
             Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1] else
           if (AreaType=FAreaType+0) or (AreaType=FAreaType+1) then
             IUrbEnter:=PLeave*
             Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]/PEnter;
         end;
       end;
       RegionalOutMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
       RegionalOutMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] + HUrbLeave+IUrbLeave;
       if ExogPopChange<0 then
       RegionalOutMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
       RegionalOutMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] - ExogPopChange;

       RegionalInMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
       RegionalInMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] + HUrbEnter+IUrbEnter;
       if ExogPopChange>0 then
       RegionalInMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
       RegionalInMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] + ExogPopChange;


       {Set new cell population by applying all the demographic rates}
        Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
        Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]
         - AgeingOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {subtract ageing}
         - DeathsOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {subtract deaths}
         - MarriagesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {subtract marriages}
         - DivorcesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {subtract divorces}
         - FirstChildOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {subtract full nest}
         - EmptyNestOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {subtract empty nest}
         - LeaveNestOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {subtract leave nest}
         - AcculturationOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {subtract acculturation}
         - WorkerStatusOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {subtract workforce out}
         - IncomeGroupOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {subtract income group out}
         - ForeignOutMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
         - DomesticOutMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
         - RegionalOutMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]

         + AgeingIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {add ageing}
         + BirthsIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {add births}
         + MarriagesIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {add marriages}
         + DivorcesIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {add divorces}
         + FirstChildIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {add full nest}
         + EmptyNestIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {add empty nest}
         + LeaveNestIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {add leave nest}
         + AcculturationIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {add acculturation}
         + WorkerStatusIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {subtract workforce out}
         + IncomeGroupIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] {subtract income group out}
         + ForeignInMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
         + DomesticInMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
         + RegionalInMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep];

   if abs(Year-testWriteYear)<0.1 then
   writeln(outest,Year:1:0,',',AreaType,',',AgeGroup,',',HhldType,',',EthnicGr,',',IncomeGr,',',WorkerGr,',',
   Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',
   Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep-1]:5:4,',',
   AgeingOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {subtract ageing}
   DeathsOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {subtract deaths}
   MarriagesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {subtract marriages}
   DivorcesOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {subtract divorces}
   FirstChildOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {subtract full nest}
   EmptyNestOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {subtract empty nest}
   LeaveNestOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {subtract leave nest}
   AcculturationOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {subtract acculturation}
   WorkerStatusOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {subtract workforce out}
   IncomeGroupOut[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {subtract income group out}
   ForeignOutMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',
   DomesticOutMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',
   RegionalOutMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',

   AgeingIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {add ageing}
   BirthsIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {add births}
   MarriagesIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {add marriages}
   DivorcesIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {add divorces}
   FirstChildIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {add full nest}
   EmptyNestIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {add empty nest}
   LeaveNestIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {add leave nest}
   AcculturationIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {add acculturation}
   WorkerStatusIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {subtract workforce out}
   IncomeGroupIn[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',   {subtract income group out}
   ForeignInMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',
   DomesticInMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',
   RegionalInMigration[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:5:4,',',

   ExogPopChange:5:4,',',
   HUrbEnter:5:4,',',
   HUrbLeave:5:4,',',
   IUrbEnter:5:4,',',
   IUrbLeave:5:4);

      end;

end; {ApplyDemographicTransitionRates}


procedure ApplyEmploymentTransitionRates(timeStep:integer);
var
{indices}
AreaType,
EmploymentType: byte;
JGrowth,ExogJobChange:single;
begin


      {apply transition in number of workers for each cell}
      for AreaType:=1 to NumberOfAreaTypes do
      for EmploymentType:=1 to NumberOfEmploymentTypes do begin

        {adjust job growth rate first for exogenous change}
        if EmploymentType=1 then begin
          if AreaType=1 then JGrowth:=ExogenousEmploymentChangeRate1A[timeStep]-ExogenousEmploymentChangeRate1A[timeStep-1] else
          if AreaType=2 then JGrowth:=ExogenousEmploymentChangeRate2A[timeStep]-ExogenousEmploymentChangeRate2A[timeStep-1] else
          if AreaType=3 then JGrowth:=ExogenousEmploymentChangeRate3A[timeStep]-ExogenousEmploymentChangeRate3A[timeStep-1] else
          if AreaType=4 then JGrowth:=ExogenousEmploymentChangeRate4A[timeStep]-ExogenousEmploymentChangeRate4A[timeStep-1] else
          if AreaType=5 then JGrowth:=ExogenousEmploymentChangeRate5A[timeStep]-ExogenousEmploymentChangeRate5A[timeStep-1] else
          if AreaType=6 then JGrowth:=ExogenousEmploymentChangeRate6A[timeStep]-ExogenousEmploymentChangeRate6A[timeStep-1] else
          if AreaType=7 then JGrowth:=ExogenousEmploymentChangeRate7A[timeStep]-ExogenousEmploymentChangeRate7A[timeStep-1] else
          if AreaType=8 then JGrowth:=ExogenousEmploymentChangeRate8A[timeStep]-ExogenousEmploymentChangeRate8A[timeStep-1] else
          if AreaType=9 then JGrowth:=ExogenousEmploymentChangeRate9A[timeStep]-ExogenousEmploymentChangeRate9A[timeStep-1] else
          if AreaType=10 then JGrowth:=ExogenousEmploymentChangeRate10A[timeStep]-ExogenousEmploymentChangeRate10A[timeStep-1] else
          if AreaType=11 then JGrowth:=ExogenousEmploymentChangeRate11A[timeStep]-ExogenousEmploymentChangeRate11A[timeStep-1] else
          if AreaType=12 then JGrowth:=ExogenousEmploymentChangeRate12A[timeStep]-ExogenousEmploymentChangeRate12A[timeStep-1];
        end else
        if EmploymentType=2 then begin
          if AreaType=1 then JGrowth:=ExogenousEmploymentChangeRate1B[timeStep]-ExogenousEmploymentChangeRate1B[timeStep-1] else
          if AreaType=2 then JGrowth:=ExogenousEmploymentChangeRate2B[timeStep]-ExogenousEmploymentChangeRate2B[timeStep-1] else
          if AreaType=3 then JGrowth:=ExogenousEmploymentChangeRate3B[timeStep]-ExogenousEmploymentChangeRate3B[timeStep-1] else
          if AreaType=4 then JGrowth:=ExogenousEmploymentChangeRate4B[timeStep]-ExogenousEmploymentChangeRate4B[timeStep-1] else
          if AreaType=5 then JGrowth:=ExogenousEmploymentChangeRate5B[timeStep]-ExogenousEmploymentChangeRate5B[timeStep-1] else
          if AreaType=6 then JGrowth:=ExogenousEmploymentChangeRate6B[timeStep]-ExogenousEmploymentChangeRate6B[timeStep-1] else
          if AreaType=7 then JGrowth:=ExogenousEmploymentChangeRate7B[timeStep]-ExogenousEmploymentChangeRate7B[timeStep-1] else
          if AreaType=8 then JGrowth:=ExogenousEmploymentChangeRate8B[timeStep]-ExogenousEmploymentChangeRate8B[timeStep-1] else
          if AreaType=9 then JGrowth:=ExogenousEmploymentChangeRate9B[timeStep]-ExogenousEmploymentChangeRate9B[timeStep-1] else
          if AreaType=10 then JGrowth:=ExogenousEmploymentChangeRate10B[timeStep]-ExogenousEmploymentChangeRate10B[timeStep-1] else
          if AreaType=11 then JGrowth:=ExogenousEmploymentChangeRate11B[timeStep]-ExogenousEmploymentChangeRate11B[timeStep-1] else
          if AreaType=12 then JGrowth:=ExogenousEmploymentChangeRate12B[timeStep]-ExogenousEmploymentChangeRate12B[timeStep-1];
        end else
        if EmploymentType=3 then begin
          if AreaType=1 then JGrowth:=ExogenousEmploymentChangeRate1C[timeStep]-ExogenousEmploymentChangeRate1C[timeStep-1] else
          if AreaType=2 then JGrowth:=ExogenousEmploymentChangeRate2C[timeStep]-ExogenousEmploymentChangeRate2C[timeStep-1] else
          if AreaType=3 then JGrowth:=ExogenousEmploymentChangeRate3C[timeStep]-ExogenousEmploymentChangeRate3C[timeStep-1] else
          if AreaType=4 then JGrowth:=ExogenousEmploymentChangeRate4C[timeStep]-ExogenousEmploymentChangeRate4C[timeStep-1] else
          if AreaType=5 then JGrowth:=ExogenousEmploymentChangeRate5C[timeStep]-ExogenousEmploymentChangeRate5C[timeStep-1] else
          if AreaType=6 then JGrowth:=ExogenousEmploymentChangeRate6C[timeStep]-ExogenousEmploymentChangeRate6C[timeStep-1] else
          if AreaType=7 then JGrowth:=ExogenousEmploymentChangeRate7C[timeStep]-ExogenousEmploymentChangeRate7C[timeStep-1] else
          if AreaType=8 then JGrowth:=ExogenousEmploymentChangeRate8C[timeStep]-ExogenousEmploymentChangeRate8C[timeStep-1] else
          if AreaType=9 then JGrowth:=ExogenousEmploymentChangeRate9C[timeStep]-ExogenousEmploymentChangeRate9C[timeStep-1] else
          if AreaType=10 then JGrowth:=ExogenousEmploymentChangeRate10C[timeStep]-ExogenousEmploymentChangeRate10C[timeStep-1] else
          if AreaType=11 then JGrowth:=ExogenousEmploymentChangeRate11C[timeStep]-ExogenousEmploymentChangeRate11C[timeStep-1] else
          if AreaType=12 then JGrowth:=ExogenousEmploymentChangeRate12C[timeStep]-ExogenousEmploymentChangeRate12C[timeStep-1];
        end;
        ExogJobChange:= JGrowth*Jobs[AreaType][EmploymentType][0];

       {Set new cell population by applying all the demographic rates}
        Jobs[AreaType][EmploymentType][timeStep]:=
         Jobs[AreaType][EmploymentType][timeStep-1]
        +JobsCreated[AreaType][EmploymentType][timeStep]
        -JobsLost[AreaType][EmploymentType][timeStep]
        +JobsMovedIn[AreaType][EmploymentType][timeStep]
        -JobsMovedOut[AreaType][EmploymentType][timeStep]
        +ExogJobChange;
      end;

end; {ApplyEmploymentTransitionRates}

procedure ApplyLandUseTransitionRates(timeStep:integer);
var
{indices}
AreaType, LandUseType: byte;
begin

     {apply transition rates for all land use types}
      for AreaType:=1 to NumberOfAreaTypes do
      for LandUseType:=1 to NumberOfLandUseTypes do begin

        Land[AreaType][LandUseType][timeStep]:=
         Land[AreaType][LandUseType][timeStep-1]
        +ChangeInLandUseIn[AreaType][LandUseType][timeStep]
        -ChangeInLandUseOut[AreaType][LandUseType][timeStep]

      end;

end; {ApplyLandUseTransitionRates}

procedure ApplyTransportationSupplyTransitionRates(timeStep:integer);
var
{indices}
AreaType,
RoadType,TransitType: byte;
begin

      for AreaType:=1 to NumberOfAreaTypes do begin

         for RoadType:=1 to NumberOfRoadTypes do
          RoadLaneMiles[AreaType][RoadType][timeStep]:=
          RoadLaneMiles[AreaType][RoadType][timeStep-1]
         +RoadLaneMilesAdded[AreaType][RoadType][timeStep]
         -RoadLaneMilesLost[AreaType][RoadType][timeStep];

         for TransitType:=1 to NumberOfTransitTypes do
          TransitRouteMiles[AreaType][TransitType][timeStep]:=
          TransitRouteMiles[AreaType][TransitType][timeStep-1]
         +TransitRouteMilesAdded[AreaType][TransitType][timeStep]
         -TransitRouteMilesLost[AreaType][TransitType][timeStep];
      end;
end; {ApplyTransportationSuppplyTransitionRates}

procedure CalculateTravelDemand(timeStep:integer);
var
{indices}
AreaType,
WorkerGr,
IncomeGr,
EthnicGr,
HhldType,
AgeGroup,
CarOwnershipLevel,
TripPurpose: byte;

FullCarUtility,CarCompUtility,NoCarUtility,
CarDriverUtility,CarPassengerUtility,TransitUtility,WalkBikeUtility,
tempCarDriverTrips,tempCarPassengerTrips,tempTransitTrips,tempWalkBikeTrips,
tempCarDriverMiles,tempCarPassengerMiles,tempTransitMiles,tempWalkBikeMiles,
tempTrips,tempPop,prob,ageBorn:single;

VarValue:array[1..NumberOfTravelModelVariables] of single;


function TravelModelEquationResult(modelNumber,firstVar,lastVar:integer):single;
var value:single; varNumber:integer;
begin
      value:=0;
      for varNumber:=firstVar to lastVar do begin
        value:=value + + TravelModelParameter[modelNumber,varNumber] * VarValue[varNumber];
      end;
      TravelModelEquationResult := value;
end;

begin


      {apply travel demand models for each cell}
      for AreaType:=1 to NumberOfAreaTypes do
      for WorkerGr:=1 to NumberOfWorkerGrs do
      for IncomeGr:=1 to NumberOfIncomeGrs do
      for EthnicGr:=1 to NumberOfEthnicGrs do
      for HhldType:=1 to NumberOfHhldTypes do
      for AgeGroup:=1 to NumberOfAgeGroups do begin

        {initialize}
        WorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=0;
        NonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=0;

        CarDriverWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=0;
        CarPassengerWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=0;
        TransitWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=0;
        WalkBikeWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        CarDriverWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=  0;
        CarPassengerWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        TransitWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=0;

        CarDriverNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=0;
        CarPassengerNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=0;
        TransitNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=0;
        WalkBikeNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        CarDriverNonWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=  0;
        CarPassengerNonWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:= 0;
        TransitNonWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=0;

{ set the initial array of input variables }
        VarValue[1]:= 1.0; {constant}
        VarValue[2]:= 0.0; {1995}
        VarValue[3]:= 0.0; {2001}
        VarValue[4]:= max(0.9 - (timeStep*TimeStepLength*0.15) , 0) ; {2009 vs 2016}

        VarValue[5]:= Dummy(AgeGroup,1);
        VarValue[6]:= Dummy(AgeGroup,2);
        VarValue[7]:= Dummy(AgeGroup,3);
        VarValue[8]:= Dummy(AgeGroup,4);
        VarValue[9]:= Dummy(AgeGroup,5);
        VarValue[10]:= Dummy(AgeGroup,6);
        VarValue[11]:= Dummy(AgeGroup,7);
        VarValue[12]:= Dummy(AgeGroup,8);
        VarValue[13]:= Dummy(AgeGroup,9);
        VarValue[14]:= Dummy(AgeGroup,10);
        VarValue[15]:= Dummy(AgeGroup,11);
        VarValue[16]:= Dummy(AgeGroup,12);
        VarValue[17]:= Dummy(AgeGroup,13);
        VarValue[18]:= Dummy(AgeGroup,14);
        VarValue[19]:= Dummy(AgeGroup,15);
        VarValue[20]:= Dummy(AgeGroup,16);
        VarValue[21]:= Dummy(AgeGroup,17);

        ageBorn:=StartYear+timeStep*TimeStepLength - 5*AgeGroup + 2.5;
        VarValue[22]:= DummyRange(ageBorn,1900,1925)*ExogenousEffectOnAgeCohortVariables[timeStep];
        VarValue[23]:= DummyRange(ageBorn,1925,1935)*ExogenousEffectOnAgeCohortVariables[timeStep];
        VarValue[24]:= DummyRange(ageBorn,1935,1945)*ExogenousEffectOnAgeCohortVariables[timeStep];
        VarValue[25]:= DummyRange(ageBorn,1945,1955)*ExogenousEffectOnAgeCohortVariables[timeStep];
        VarValue[26]:= DummyRange(ageBorn,1955,1965)*ExogenousEffectOnAgeCohortVariables[timeStep];
        VarValue[27]:= DummyRange(ageBorn,1965,1975)*ExogenousEffectOnAgeCohortVariables[timeStep];
        VarValue[28]:= DummyRange(ageBorn,1975,1985)*ExogenousEffectOnAgeCohortVariables[timeStep];
        VarValue[29]:= DummyRange(ageBorn,1985,1995)*ExogenousEffectOnAgeCohortVariables[timeStep];
        VarValue[30]:= DummyRange(ageBorn,1995,2005)*ExogenousEffectOnAgeCohortVariables[timeStep];
        VarValue[31]:= DummyRange(ageBorn,2005,2075)*ExogenousEffectOnAgeCohortVariables[timeStep];

        VarValue[32]:= Dummy(HhldType,2) + Dummy(HhldType,4);
        VarValue[33]:= Dummy(HhldType,3) + Dummy(HhldType,4);
        VarValue[34]:= Dummy(HhldType,3);

        VarValue[35]:= Dummy(EthnicGr,1) + Dummy(EthnicGr,2) + Dummy(EthnicGr,3);
        VarValue[36]:= Dummy(EthnicGr,4) + Dummy(EthnicGr,5) + Dummy(EthnicGr,6);
        VarValue[37]:= Dummy(EthnicGr,7) + Dummy(EthnicGr,8) + Dummy(EthnicGr,9);
        VarValue[38]:= 1.0 - (Dummy(EthnicGr,1) + Dummy(EthnicGr,4) + Dummy(EthnicGr,7) + Dummy(EthnicGr,10));
        VarValue[39]:= Dummy(EthnicGr,3) + Dummy(EthnicGr,6) + Dummy(EthnicGr,9) + Dummy(EthnicGr,12);

        VarValue[40]:= Dummy(WorkerGr,1);

        VarValue[41]:= Dummy(IncomeGr,1);
        VarValue[42]:= Dummy(IncomeGr,3);

        VarValue[43]:= Dummy(AreaTypeDensity[AreaType],1);
        VarValue[44]:= Dummy(AreaTypeDensity[AreaType],2);
        VarValue[45]:= Dummy(AreaTypeDensity[AreaType],3);
        VarValue[46]:= Dummy(AreaTypeDensity[AreaType],4);
        VarValue[47]:= Dummy(AreaTypeDensity[AreaType],5);

        VarValue[48]:= 1.0; {rail MSA}
        VarValue[49]:= 1.0; {large MSA}
        VarValue[50]:= 1.0; {DVRPC MSA}

        VarValue[51]:= 0.0; {gas price not used in first model}
        VarValue[52]:= 0.0; {proxy}
        VarValue[53]:= 0.0; {no diary}

       {apply the auto ownership model}

        FullCarUtility:= 1.0; {base}
        CarCompUtility:= exp(TravelModelEquationResult(CarOwnership_CarCompetition,1,53))
          * ExogenousEffectOnSharedCarFraction[timeStep];
        NoCarUtility:= exp(TravelModelEquationResult(CarOwnership_NoCar,1,53))
          * ExogenousEffectOnNoCarFraction[timeStep];

        {apply the rest of the models conditional on car ownership}
        for CarOwnershipLevel:= 1 to 3 do begin

          if CarOwnershipLevel = 1 then begin
            prob:=FullCarUtility / (FullCarUtility + CarCompUtility + NoCarUtility);
            OwnCar[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
              Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] * prob;
          end else
          if CarOwnershipLevel = 2 then begin
            prob:=CarCompUtility / (FullCarUtility + CarCompUtility + NoCarUtility);
            ShareCar[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
              Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] * prob;
          end else
            prob:=NoCarUtility / (FullCarUtility + CarCompUtility + NoCarUtility);
            NoCar[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
              Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] * prob;

          tempPop:=Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] * prob;

          VarValue[54]:= Dummy(CarOwnershipLevel,3);
          VarValue[55]:= Dummy(CarOwnershipLevel,2);
          VarValue[51]:= BaseGasolinePrice * ExogenousEffectOnGasolinePrice[timeStep];

          {loop on trip purposes 1= work, 2= non-work, 3=child }
          for TripPurpose:=1 to 3 do begin

             VarValue[56]:= Dummy(TripPurpose,1);

            {apply trip generation model, work trips only for workers}
            if (TripPurpose=1) and (WorkerGr = 1) and (AgeGroup>3) then begin
              tempTrips:= tempPop * (exp(TravelModelEquationResult(WorkTrip_Generation,1,56)) - 1.0)
              * ExogenousEffectOnWorkTripRate[timeStep];
              WorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                WorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] + tempTrips;
            end
            else if (TripPurpose=2) and (AgeGroup>3) then begin
              tempTrips:= tempPop * (exp(TravelModelEquationResult(NonWorkTrip_Generation,1,56)) - 1.0)
              * ExogenousEffectOnNonWorkTripRate[timeStep];
              NonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                NonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] + tempTrips;
            end
            else if (TripPurpose=3) and (AgeGroup<=3) then begin
               tempTrips:= tempPop * (exp(TravelModelEquationResult(ChildTrip_Generation,1,56)) - 1.0)
               * ExogenousEffectOnNonWorkTripRate[timeStep];
               NonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                 NonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep] + tempTrips;
             end
            else tempTrips:=0;

            if tempTrips>0 then begin

              {set mode utilities}
              if (TripPurpose=1) and (WorkerGr = 1) and (AgeGroup>3) then begin
                CarDriverUtility:= 1.0; {base}
                CarPassengerUtility:= exp(TravelModelEquationResult(WorkTrip_CarPassengerMode,1,56))
                  * ExogenousEffectOnCarPassengerModeFraction[timeStep];
                TransitUtility:= exp(TravelModelEquationResult(WorkTrip_TransitMode,1,56))
                 * ExogenousEffectOnTransitModeFraction[timeStep];
                WalkBikeUtility:= exp(TravelModelEquationResult(WorkTrip_WalkBikeMode,1,56))
                 * ExogenousEffectOnWalkBikeModeFraction[timeStep];
              end
              else if (TripPurpose=2) and (AgeGroup>3) then begin
                CarDriverUtility:= 1.0; {base}
                CarPassengerUtility:= exp(TravelModelEquationResult(NonWorkTrip_CarPassengerMode,1,56))
                 * ExogenousEffectOnCarPassengerModeFraction[timeStep];
                TransitUtility:= exp(TravelModelEquationResult(NonWorkTrip_TransitMode,1,56))
                 * ExogenousEffectOnTransitModeFraction[timeStep];
                WalkBikeUtility:= exp(TravelModelEquationResult(NonWorkTrip_WalkBikeMode,1,56))
                * ExogenousEffectOnWalkBikeModeFraction[timeStep];
              end
              else if (TripPurpose=3) and (AgeGroup<=3) then begin
                CarDriverUtility:= 1.0; {school bus for kids}
                CarPassengerUtility:=exp(TravelModelEquationResult(ChildTrip_CarPassengerMode,1,56))
                 * ExogenousEffectOnCarPassengerModeFraction[timeStep];
                TransitUtility:= exp(TravelModelEquationResult(ChildTrip_TransitMode,1,56))
                 * ExogenousEffectOnTransitModeFraction[timeStep];
                WalkBikeUtility:= exp(TravelModelEquationResult(ChildTrip_WalkBikeMode,1,56))
                * ExogenousEffectOnWalkBikeModeFraction[timeStep];
              end;
              {split trips by mode and apply distance models}
              if (TripPurpose = 3) then tempCarDriverTrips := 0 else
              tempCarDriverTrips := tempTrips *
                CarDriverUtility / (CarDriverUtility + CarPassengerUtility + TransitUtility + WalkBikeUtility);

              tempCarDriverMiles := tempCarDriverTrips *
               (exp(TravelModelEquationResult(CarDriverTrip_Distance, 1,56)) - 1.0)
               * ExogenousEffectOnCarTripDistance[timeStep];

              tempCarPassengerTrips := tempTrips *
                CarPassengerUtility / (CarDriverUtility + CarPassengerUtility + TransitUtility + WalkBikeUtility);

              tempCarPassengerMiles := tempCarPassengerTrips *
               (exp(TravelModelEquationResult(CarPassengerTrip_Distance, 1,56)) - 1.0)
               * ExogenousEffectOnCarTripDistance[timeStep];

              tempTransitTrips := tempTrips *
                TransitUtility / (CarDriverUtility + CarPassengerUtility + TransitUtility + WalkBikeUtility);

              tempTransitMiles := tempTransitTrips *
               (exp(TravelModelEquationResult(TransitTrip_Distance, 1,56)) - 1.0);

              tempWalkBikeTrips := tempTrips *
                WalkBikeUtility / (CarDriverUtility + CarPassengerUtility + TransitUtility + WalkBikeUtility);


              if (TripPurpose = 1) then begin
                CarDriverWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                CarDriverWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempCarDriverTrips;
                CarPassengerWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                CarPassengerWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempCarPassengerTrips;
                TransitWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                TransitWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempTransitTrips;
                WalkBikeWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                WalkBikeWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempWalkBikeTrips;

                CarDriverWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                CarDriverWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempCarDriverMiles;
                CarPassengerWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                CarPassengerWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempCarPassengerMiles;
                TransitWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                TransitWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempTransitMiles;
              end else begin
                CarDriverNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                CarDriverNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempCarDriverTrips;
                CarPassengerNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                CarPassengerNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempCarPassengerTrips;
                TransitNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                TransitNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempTransitTrips;
                WalkBikeNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                WalkBikeNonWorkTrips[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempWalkBikeTrips;

                CarDriverNonWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                CarDriverNonWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempCarDriverMiles;
                CarPassengerNonWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                CarPassengerNonWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempCarPassengerMiles;
                TransitNonWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]:=
                TransitNonWorkMiles[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][timeStep]
                + tempTransitMiles;
              end;

            end; {trips for purpose}
          end;  {purpose loop}
        end; {car ownership loop}
      end; {cells}

end; {CalculateTravelDemand}

procedure writeSimulationResults;
var ouf:text; ts,demVar:integer;
{indices}
AreaType,
WorkerGr,
IncomeGr,
EthnicGr,
HhldType,
AgeGroup,
EmploymentType,
LandUseType,
RoadType,
TransitType: byte;

procedure writeTimeArray(demArray:TimeStepArray; demLabel:string; varLabel:string);
var ts,tx:integer;
begin
  if varLabel = 'Population' then write(ouf,demLabel) else
  if demLabel = 'Total' then write(ouf,varLabel) else
  if demLabel = '' then write(ouf,varLabel) else
    write(ouf,varLabel+'-'+demLabel);
  for ts:=0 to NumberOfTimeSteps do begin
    if (ts=0) and (demArray[ts]=0) then tx:=1 else tx:=ts; {avoids 0 rate in first period}
    write(ouf,',',demArray[tx]:4:2);
  end;
  writeln(ouf);
end;

begin

   assign(ouf,OutputDirectory+RunLabel+'.csv'); rewrite(ouf);

   write(ouf,'Year');
   for ts:=0 to NumberOfTimeSteps do begin
     write(ouf,',',StartYear + ts*TimeStepLength:4:1);
   end;
   writeln(ouf);

   for demVar:=1 to NumberOfDemographicVariables do
   for AgeGroup:=0 to 0 do writeTimeArray(AgeGroupMarginals[demVar][AgeGroup],'',DemographicVariableLabels[demVar]);

   for AreaType:=1 to NumberOfAreaTypes do
   for EmploymentType:=1 to NumberOfEmploymentTypes do
      writeTimeArray(Jobs[AreaType][EmploymentType],'',
      AreaTypeLabels[AreaType]+'/'+EmploymentTypeLabels[EmploymentType]);

   for AreaType:=1 to NumberOfAreaTypes do
   for LandUseType:=1 to NumberOfLandUseTypes do
      writeTimeArray(Land[AreaType][LandUseType],'',
      AreaTypeLabels[AreaType]+'/'+LandUseTypeLabels[LandUseType]);

   for AreaType:=1 to NumberOfAreaTypes do
   for RoadType:=1 to NumberOfRoadTypes do
      writeTimeArray(RoadLaneMiles[AreaType][RoadType],
      AreaTypeLabels[AreaType]+'/'+RoadTypeLabels[RoadType],'LaneMiles');

   for AreaType:=1 to NumberOfAreaTypes do
   for TransitType:=1 to NumberOfTransitTypes do
      writeTimeArray(TransitRouteMiles[AreaType][TransitType],
      AreaTypeLabels[AreaType]+'/'+TransitTypeLabels[TransitType],'RouteMiles');

   for demVar:=1 to NumberOfDemographicVariables do begin
     for AgeGroup:=1 to NumberOfAgeGroups do writeTimeArray(AgeGroupMarginals[demVar][AgeGroup],AgeGroupLabels[AgeGroup],DemographicVariableLabels[demVar]);
     for HhldType:=1 to NumberOfHhldTypes do writeTimeArray(HhldTypeMarginals[demVar][HhldType],HhldTypeLabels[HhldType],DemographicVariableLabels[demVar]);
     for EthnicGr:=1 to NumberOfEthnicGrs do writeTimeArray(EthnicGrMarginals[demVar][EthnicGr],EthnicGrLabels[EthnicGr],DemographicVariableLabels[demVar]);
     for IncomeGr:=1 to NumberOfIncomeGrs do writeTimeArray(IncomeGrMarginals[demVar][IncomeGr],IncomeGrLabels[IncomeGr],DemographicVariableLabels[demVar]);
     for WorkerGr:=1 to NumberOfWorkerGrs do writeTimeArray(WorkerGrMarginals[demVar][WorkerGr],WorkerGrLabels[WorkerGr],DemographicVariableLabels[demVar]);
     for AreaType:=1 to NumberOfAreaTypes do writeTimeArray(AreaTypeMarginals[demVar][AreaType],AreaTypeLabels[AreaType],DemographicVariableLabels[demVar]);
   end;

   for AreaType:=1 to NumberOfAreaTypes do
       writeTimeArray(JobDemandSupplyIndex[AreaType],
      AreaTypeLabels[AreaType],'Job Demand/Supply');

   for AreaType:=1 to NumberOfAreaTypes do
       writeTimeArray(CommercialSpaceDemandSupplyIndex[AreaType],
      AreaTypeLabels[AreaType],'NonR.Space Demand/Supply');

   for AreaType:=1 to NumberOfAreaTypes do
       writeTimeArray(ResidentialSpaceDemandSupplyIndex[AreaType],
      AreaTypeLabels[AreaType],'Res.Space Demand/Supply');

   for AreaType:=1 to NumberOfAreaTypes do
       writeTimeArray(DevelopableSpaceDemandSupplyIndex[AreaType],
      AreaTypeLabels[AreaType],'Dev.Space Demand/Supply');

   for AreaType:=1 to NumberOfAreaTypes do
       writeTimeArray(RoadVehicleCapacityDemandSupplyIndex[AreaType],
      AreaTypeLabels[AreaType],'Road Miles Demand/Supply');

  close(ouf);

end;

procedure writeHouseholdConversionFile;
{
Income,Low,Low,Low,Low,Low,Low,Low,Low,Low,Low,Low,Low,Low,Low,Low,Low,Low,Low,Low,Low, 0.0000 ,Person characteristics, 0.0000 ,HH characteristics >>>, 0.0000 ,
HH size,1,1,2,2,2,2,2,2,3,3,3,3,3,3,4+,4+,4+,4+,4+,4+, 0.0000 ,   V, 0.0000 , 0.0000 , 0.0000 ,
HH workers,0,1,0,0,1,1,2+,2+,0,0,1,1,2+,2+,0,0,1,1,2+,2+, 0.0000 ,   V, 0.0000 , 0.0000 , 0.0000 ,
Children?,No,No,No,Yes,No,Yes,No,Yes,No,Yes,No,Yes,No,Yes,No,Yes,No,Yes,No,Yes, 0.0000 ,   V, 0.0000 , 0.0000 , 0.0000 ,
Code,1111,1121,1211,1212,1221,1222,1231,1232,1311,1312,1321,1322,1331,1332,1411,1412,1421,1422,1431,1432,Total,HH income,"Marital status, kids",Work status,Age group,
}
const nHInc = 3; nHComp=20; nHSize=4; nHWork=3; nHKids=2;
 HIncLabel:array[1..nHInc]   of string= ('LowInc','MedInc','Hi Inc');
 HSizeLabel:array[1..nHSize] of string= ('1 pers','2 pers','3 pers','4+pers');
 HWorkLabel:array[1..nHWork] of string= ('0 wkrs','1 wrkr','2+wkrs');
 HKidsLabel:array[1..nHKids] of string= ('0 kids','1+kids');

 HCompSize:array[1..nHComp] of byte= (1,1, 2,2,2,2,2,2, 3,3,3,3,3,3, 4,4,4,4,4,4);
 HCompWork:array[1..nHComp] of byte= (0,1, 0,0,1,1,2,2, 0,0,1,1,2,2, 0,0,1,1,2,2);
 HCompKids:array[1..nHComp] of byte= (0,0, 0,1,0,1,0,1, 0,1,0,1,0,1, 0,1,0,1,0,1);


AvgHSize:array[1..nHInc,1..nHComp] of single=((1.00,1.00,2.00,2.00,2.00,2.00,2.00,2.00,3.00,3.00,3.00,3.00,3.00,3.00,4.48,4.96,4.61,5.03,4.62,5.42),
                                              (1.00,1.00,2.00,2.00,2.00,2.00,2.00,2.00,3.00,3.00,3.00,3.00,3.00,3.00,4.35,4.76,4.29,4.78,4.39,4.87),
                                              (1.00,1.00,2.00,2.00,2.00,2.00,2.00,2.00,3.00,3.00,3.00,3.00,3.00,3.00,4.46,4.72,4.22,4.66,4.32,4.64));

HSizeAdj:array[1..nHSize] of single = (0.91,0.91,1.1,1.15);

nPInc=3; nPHType=4; nPWork=2; nPAgeG=6;

phFrac:array[1..nPInc,1..nPHType,1..nPWork,1..nPAgeG,1..nHComp] of single=
{IHWA       1p0w0k   1p1w0k   2p0w0k   2p0w+k   2p1w0k   2p1w+k   2p2w0k   2p2w+k   3p0w0k   3p0w+k   3p1w0k   3p1w+k   3p2w0k   3p2w+k   4p0w0k   4p0w+k   4p1w0k   4p1w+k   4p2w0k   4p2w+k   }
{1111}((((( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{1112}    ( 0.1946 , 0.0000 , 0.2265 , 0.0000 , 0.1549 , 0.0000 , 0.0000 , 0.0000 , 0.0768 , 0.0000 , 0.0984 , 0.0000 , 0.0503 , 0.0000 , 0.0487 , 0.0000 , 0.0571 , 0.0000 , 0.0926 , 0.0000)   ,
{1113}    ( 0.3873 , 0.0000 , 0.2825 , 0.0000 , 0.1264 , 0.0000 , 0.0000 , 0.0000 , 0.0571 , 0.0000 , 0.0591 , 0.0000 , 0.0236 , 0.0000 , 0.0269 , 0.0000 , 0.0232 , 0.0000 , 0.0139 , 0.0000)   ,
{1114}    ( 0.5826 , 0.0000 , 0.1803 , 0.0000 , 0.0881 , 0.0000 , 0.0000 , 0.0000 , 0.0466 , 0.0000 , 0.0424 , 0.0000 , 0.0194 , 0.0000 , 0.0089 , 0.0000 , 0.0118 , 0.0000 , 0.0198 , 0.0000)   ,
{1115}    ( 0.7164 , 0.0000 , 0.1382 , 0.0000 , 0.0639 , 0.0000 , 0.0000 , 0.0000 , 0.0268 , 0.0000 , 0.0243 , 0.0000 , 0.0090 , 0.0000 , 0.0072 , 0.0000 , 0.0067 , 0.0000 , 0.0076 , 0.0000)   ,
{1116}    ( 0.7439 , 0.0000 , 0.1131 , 0.0000 , 0.0754 , 0.0000 , 0.0000 , 0.0000 , 0.0235 , 0.0000 , 0.0194 , 0.0000 , 0.0061 , 0.0000 , 0.0041 , 0.0000 , 0.0040 , 0.0000 , 0.0106 , 0.0000))  ,
{1121}   (( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{1122}    ( 0.0000 , 0.2347 , 0.0000 , 0.0000 , 0.1729 , 0.0000 , 0.2054 , 0.0000 , 0.0000 , 0.0000 , 0.0369 , 0.0000 , 0.1794 , 0.0000 , 0.0000 , 0.0000 , 0.0122 , 0.0000 , 0.1584 , 0.0000)   ,
{1123}    ( 0.0000 , 0.4038 , 0.0000 , 0.0000 , 0.2199 , 0.0000 , 0.1625 , 0.0000 , 0.0000 , 0.0000 , 0.0323 , 0.0000 , 0.0994 , 0.0000 , 0.0000 , 0.0000 , 0.0076 , 0.0000 , 0.0744 , 0.0000)   ,
{1124}    ( 0.0000 , 0.5653 , 0.0000 , 0.0000 , 0.1692 , 0.0000 , 0.1266 , 0.0000 , 0.0000 , 0.0000 , 0.0323 , 0.0000 , 0.0590 , 0.0000 , 0.0000 , 0.0000 , 0.0061 , 0.0000 , 0.0415 , 0.0000)   ,
{1125}    ( 0.0000 , 0.7243 , 0.0000 , 0.0000 , 0.1046 , 0.0000 , 0.0889 , 0.0000 , 0.0000 , 0.0000 , 0.0047 , 0.0000 , 0.0513 , 0.0000 , 0.0000 , 0.0000 , 0.0058 , 0.0000 , 0.0203 , 0.0000)   ,
{1126}    ( 0.0000 , 0.7243 , 0.0000 , 0.0000 , 0.1046 , 0.0000 , 0.0889 , 0.0000 , 0.0000 , 0.0000 , 0.0047 , 0.0000 , 0.0513 , 0.0000 , 0.0000 , 0.0000 , 0.0058 , 0.0000 , 0.0203 , 0.0000))) ,
{1211}  ((( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{1212}    ( 0.0000 , 0.0000 , 0.0781 , 0.0000 , 0.2421 , 0.0000 , 0.0000 , 0.0000 , 0.1213 , 0.0000 , 0.1468 , 0.0000 , 0.1061 , 0.0000 , 0.0089 , 0.0000 , 0.1188 , 0.0000 , 0.1779 , 0.0000)   ,
{1213}    ( 0.0000 , 0.0000 , 0.3107 , 0.0000 , 0.2975 , 0.0000 , 0.0000 , 0.0000 , 0.1025 , 0.0000 , 0.0822 , 0.0000 , 0.0477 , 0.0000 , 0.0223 , 0.0000 , 0.0782 , 0.0000 , 0.0589 , 0.0000)   ,
{1214}    ( 0.0000 , 0.0000 , 0.3486 , 0.0000 , 0.2790 , 0.0000 , 0.0000 , 0.0000 , 0.0884 , 0.0000 , 0.1003 , 0.0000 , 0.0698 , 0.0000 , 0.0325 , 0.0000 , 0.0374 , 0.0000 , 0.0440 , 0.0000)   ,
{1215}    ( 0.0000 , 0.0000 , 0.6490 , 0.0000 , 0.1662 , 0.0000 , 0.0000 , 0.0000 , 0.0452 , 0.0000 , 0.0730 , 0.0000 , 0.0176 , 0.0000 , 0.0161 , 0.0000 , 0.0126 , 0.0000 , 0.0203 , 0.0000)   ,
{1216}    ( 0.0000 , 0.0000 , 0.8346 , 0.0000 , 0.0333 , 0.0000 , 0.0000 , 0.0000 , 0.0513 , 0.0000 , 0.0546 , 0.0000 , 0.0056 , 0.0000 , 0.0027 , 0.0000 , 0.0103 , 0.0000 , 0.0076 , 0.0000))  ,
{1221}   (( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{1222}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0933 , 0.0000 , 0.3741 , 0.0000 , 0.0000 , 0.0000 , 0.0669 , 0.0000 , 0.2510 , 0.0000 , 0.0000 , 0.0000 , 0.0146 , 0.0000 , 0.2000 , 0.0000)   ,
{1223}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.1532 , 0.0000 , 0.4193 , 0.0000 , 0.0000 , 0.0000 , 0.1467 , 0.0000 , 0.1418 , 0.0000 , 0.0000 , 0.0000 , 0.0214 , 0.0000 , 0.1177 , 0.0000)   ,
{1224}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.2715 , 0.0000 , 0.3783 , 0.0000 , 0.0000 , 0.0000 , 0.0686 , 0.0000 , 0.1736 , 0.0000 , 0.0000 , 0.0000 , 0.0185 , 0.0000 , 0.0895 , 0.0000)   ,
{1225}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.4895 , 0.0000 , 0.3094 , 0.0000 , 0.0000 , 0.0000 , 0.0364 , 0.0000 , 0.1194 , 0.0000 , 0.0000 , 0.0000 , 0.0057 , 0.0000 , 0.0395 , 0.0000)   ,
{1226}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.4895 , 0.0000 , 0.3094 , 0.0000 , 0.0000 , 0.0000 , 0.0364 , 0.0000 , 0.1194 , 0.0000 , 0.0000 , 0.0000 , 0.0057 , 0.0000 , 0.0395 , 0.0000))) ,
{1311}  ((( 0.0000 , 0.0000 , 0.0000 , 0.0362 , 0.0000 , 0.0706 , 0.0000 , 0.0000 , 0.0000 , 0.0663 , 0.0000 , 0.1740 , 0.0000 , 0.0299 , 0.0000 , 0.1343 , 0.0000 , 0.3236 , 0.0000 , 0.1651)   ,
{1312}    ( 0.0000 , 0.0000 , 0.0098 , 0.0574 , 0.0000 , 0.0541 , 0.0000 , 0.0000 , 0.0372 , 0.0932 , 0.0165 , 0.1446 , 0.0000 , 0.0252 , 0.0253 , 0.1408 , 0.0416 , 0.1948 , 0.0071 , 0.1522)   ,
{1313}    ( 0.0000 , 0.0000 , 0.0120 , 0.1027 , 0.0000 , 0.0062 , 0.0000 , 0.0000 , 0.0920 , 0.1535 , 0.0037 , 0.0396 , 0.0000 , 0.0004 , 0.1052 , 0.2632 , 0.0318 , 0.1188 , 0.0070 , 0.0639)   ,
{1314}    ( 0.0000 , 0.0000 , 0.0153 , 0.1607 , 0.0000 , 0.0191 , 0.0000 , 0.0000 , 0.0306 , 0.1320 , 0.0536 , 0.0861 , 0.0000 , 0.0038 , 0.0306 , 0.1518 , 0.0542 , 0.1492 , 0.0019 , 0.1110)   ,
{1315}    ( 0.0000 , 0.0000 , 0.0117 , 0.1301 , 0.0000 , 0.0108 , 0.0000 , 0.0000 , 0.0528 , 0.0851 , 0.0656 , 0.0900 , 0.0000 , 0.0029 , 0.0763 , 0.1037 , 0.0548 , 0.1673 , 0.0166 , 0.1321)   ,
{1316}    ( 0.0000 , 0.0000 , 0.0096 , 0.0669 , 0.0000 , 0.0032 , 0.0000 , 0.0000 , 0.0478 , 0.0541 , 0.0223 , 0.0828 , 0.0000 , 0.0064 , 0.0446 , 0.1210 , 0.2070 , 0.2038 , 0.0127 , 0.1178))  ,
{1321}   (( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{1322}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0313 , 0.0439 , 0.0000 , 0.0105 , 0.0000 , 0.0000 , 0.0642 , 0.0859 , 0.0215 , 0.1053 , 0.0000 , 0.0000 , 0.0627 , 0.1487 , 0.0968 , 0.3292)   ,
{1323}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0288 , 0.1013 , 0.0000 , 0.0031 , 0.0000 , 0.0000 , 0.0445 , 0.1914 , 0.0085 , 0.0643 , 0.0000 , 0.0000 , 0.0730 , 0.1960 , 0.0545 , 0.2345)   ,
{1324}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0142 , 0.1338 , 0.0000 , 0.0097 , 0.0000 , 0.0000 , 0.0564 , 0.1383 , 0.0162 , 0.1050 , 0.0000 , 0.0000 , 0.0430 , 0.1273 , 0.1042 , 0.2518)   ,
{1325}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0070 , 0.1174 , 0.0000 , 0.0094 , 0.0000 , 0.0000 , 0.0540 , 0.1573 , 0.0094 , 0.1009 , 0.0000 , 0.0000 , 0.1127 , 0.0376 , 0.0446 , 0.3498)   ,
{1326}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0070 , 0.1174 , 0.0000 , 0.0094 , 0.0000 , 0.0000 , 0.0540 , 0.1573 , 0.0094 , 0.1009 , 0.0000 , 0.0000 , 0.1127 , 0.0376 , 0.0446 , 0.3498))) ,
{1411}  ((( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0103 , 0.0000 , 0.0512 , 0.0000 , 0.0348 , 0.0000 , 0.0619 , 0.0000 , 0.4295 , 0.0000 , 0.4123)   ,
{1412}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0030 , 0.0148 , 0.0844 , 0.0679 , 0.0000 , 0.0289 , 0.0116 , 0.0850 , 0.0750 , 0.3150 , 0.0222 , 0.2923)   ,
{1413}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0136 , 0.0206 , 0.0713 , 0.0734 , 0.0000 , 0.0016 , 0.0272 , 0.1471 , 0.0510 , 0.4816 , 0.0136 , 0.0990)   ,
{1414}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0161 , 0.0739 , 0.0067 , 0.1766 , 0.0000 , 0.0094 , 0.0168 , 0.1269 , 0.0396 , 0.2794 , 0.0302 , 0.2243)   ,
{1415}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0340 , 0.1089 , 0.0328 , 0.1113 , 0.0000 , 0.0023 , 0.0351 , 0.1616 , 0.0585 , 0.2482 , 0.0351 , 0.1721)   ,
{1416}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0038 , 0.1107 , 0.0115 , 0.0307 , 0.0000 , 0.0000 , 0.0191 , 0.1145 , 0.0420 , 0.1947 , 0.0649 , 0.4084))  ,
{1421}   (( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{1422}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0786 , 0.0277 , 0.0751 , 0.0534 , 0.0000 , 0.0000 , 0.0684 , 0.1138 , 0.1483 , 0.4346)   ,
{1423}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0391 , 0.0379 , 0.0334 , 0.0714 , 0.0000 , 0.0000 , 0.0391 , 0.2398 , 0.0652 , 0.4740)   ,
{1424}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0377 , 0.1048 , 0.0176 , 0.1429 , 0.0000 , 0.0000 , 0.0239 , 0.1806 , 0.0356 , 0.4567)   ,
{1425}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0157 , 0.1289 , 0.0189 , 0.0660 , 0.0000 , 0.0000 , 0.0346 , 0.1038 , 0.1415 , 0.4906)   ,
{1426}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0157 , 0.1289 , 0.0189 , 0.0660 , 0.0000 , 0.0000 , 0.0346 , 0.1038 , 0.1415 , 0.4906)))),
{2111} (((( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{2112}    ( 0.0451 , 0.0000 , 0.0860 , 0.0000 , 0.2547 , 0.0000 , 0.0000 , 0.0000 , 0.0168 , 0.0000 , 0.1630 , 0.0000 , 0.1504 , 0.0000 , 0.0503 , 0.0000 , 0.0645 , 0.0000 , 0.1693 , 0.0000)   ,
{2113}    ( 0.1564 , 0.0000 , 0.1818 , 0.0000 , 0.1913 , 0.0000 , 0.0000 , 0.0000 , 0.0423 , 0.0000 , 0.1015 , 0.0000 , 0.1268 , 0.0000 , 0.0359 , 0.0000 , 0.1068 , 0.0000 , 0.0571 , 0.0000)   ,
{2114}    ( 0.3200 , 0.0000 , 0.2297 , 0.0000 , 0.1597 , 0.0000 , 0.0000 , 0.0000 , 0.0365 , 0.0000 , 0.0712 , 0.0000 , 0.0856 , 0.0000 , 0.0081 , 0.0000 , 0.0307 , 0.0000 , 0.0584 , 0.0000)   ,
{2115}    ( 0.5582 , 0.0000 , 0.1701 , 0.0000 , 0.1170 , 0.0000 , 0.0000 , 0.0000 , 0.0160 , 0.0000 , 0.0560 , 0.0000 , 0.0312 , 0.0000 , 0.0046 , 0.0000 , 0.0206 , 0.0000 , 0.0262 , 0.0000)   ,
{2116}    ( 0.6034 , 0.0000 , 0.1114 , 0.0000 , 0.1273 , 0.0000 , 0.0000 , 0.0000 , 0.0141 , 0.0000 , 0.0713 , 0.0000 , 0.0317 , 0.0000 , 0.0039 , 0.0000 , 0.0166 , 0.0000 , 0.0204 , 0.0000))  ,
{2121}   (( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{2122}    ( 0.0000 , 0.2112 , 0.0000 , 0.0000 , 0.0973 , 0.0000 , 0.2974 , 0.0000 , 0.0000 , 0.0000 , 0.0106 , 0.0000 , 0.2269 , 0.0000 , 0.0000 , 0.0000 , 0.0090 , 0.0000 , 0.1477 , 0.0000)   ,
{2123}    ( 0.0000 , 0.5261 , 0.0000 , 0.0000 , 0.1290 , 0.0000 , 0.1903 , 0.0000 , 0.0000 , 0.0000 , 0.0169 , 0.0000 , 0.0883 , 0.0000 , 0.0000 , 0.0000 , 0.0067 , 0.0000 , 0.0427 , 0.0000)   ,
{2124}    ( 0.0000 , 0.5778 , 0.0000 , 0.0000 , 0.1379 , 0.0000 , 0.1365 , 0.0000 , 0.0000 , 0.0000 , 0.0214 , 0.0000 , 0.0811 , 0.0000 , 0.0000 , 0.0000 , 0.0036 , 0.0000 , 0.0418 , 0.0000)   ,
{2125}    ( 0.0000 , 0.6620 , 0.0000 , 0.0000 , 0.1365 , 0.0000 , 0.1023 , 0.0000 , 0.0000 , 0.0000 , 0.0163 , 0.0000 , 0.0635 , 0.0000 , 0.0000 , 0.0000 , 0.0016 , 0.0000 , 0.0178 , 0.0000)   ,
{2126}    ( 0.0000 , 0.6620 , 0.0000 , 0.0000 , 0.1365 , 0.0000 , 0.1023 , 0.0000 , 0.0000 , 0.0000 , 0.0163 , 0.0000 , 0.0635 , 0.0000 , 0.0000 , 0.0000 , 0.0016 , 0.0000 , 0.0178 , 0.0000))) ,
{2211}  ((( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{2212}    ( 0.0000 , 0.0000 , 0.0258 , 0.0000 , 0.1921 , 0.0000 , 0.0000 , 0.0000 , 0.0556 , 0.0000 , 0.1536 , 0.0000 , 0.2565 , 0.0000 , 0.0163 , 0.0000 , 0.0474 , 0.0000 , 0.2527 , 0.0000)   ,
{2213}    ( 0.0000 , 0.0000 , 0.0532 , 0.0000 , 0.4953 , 0.0000 , 0.0000 , 0.0000 , 0.0719 , 0.0000 , 0.1474 , 0.0000 , 0.0884 , 0.0000 , 0.0129 , 0.0000 , 0.0503 , 0.0000 , 0.0805 , 0.0000)   ,
{2214}    ( 0.0000 , 0.0000 , 0.2198 , 0.0000 , 0.4390 , 0.0000 , 0.0000 , 0.0000 , 0.0534 , 0.0000 , 0.0990 , 0.0000 , 0.0898 , 0.0000 , 0.0122 , 0.0000 , 0.0212 , 0.0000 , 0.0656 , 0.0000)   ,
{2215}    ( 0.0000 , 0.0000 , 0.5932 , 0.0000 , 0.2368 , 0.0000 , 0.0000 , 0.0000 , 0.0312 , 0.0000 , 0.0656 , 0.0000 , 0.0320 , 0.0000 , 0.0058 , 0.0000 , 0.0099 , 0.0000 , 0.0254 , 0.0000)   ,
{2216}    ( 0.0000 , 0.0000 , 0.7700 , 0.0000 , 0.0542 , 0.0000 , 0.0000 , 0.0000 , 0.0358 , 0.0000 , 0.0708 , 0.0000 , 0.0259 , 0.0000 , 0.0069 , 0.0000 , 0.0110 , 0.0000 , 0.0254 , 0.0000))  ,
{2221}   (( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{2222}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0409 , 0.0000 , 0.3616 , 0.0000 , 0.0000 , 0.0000 , 0.0269 , 0.0000 , 0.3229 , 0.0000 , 0.0000 , 0.0000 , 0.0062 , 0.0000 , 0.2415 , 0.0000)   ,
{2223}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0795 , 0.0000 , 0.5989 , 0.0000 , 0.0000 , 0.0000 , 0.0511 , 0.0000 , 0.1717 , 0.0000 , 0.0000 , 0.0000 , 0.0085 , 0.0000 , 0.0903 , 0.0000)   ,
{2224}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.1593 , 0.0000 , 0.5154 , 0.0000 , 0.0000 , 0.0000 , 0.0308 , 0.0000 , 0.2071 , 0.0000 , 0.0000 , 0.0000 , 0.0037 , 0.0000 , 0.0838 , 0.0000)   ,
{2225}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.3781 , 0.0000 , 0.4237 , 0.0000 , 0.0000 , 0.0000 , 0.0277 , 0.0000 , 0.1288 , 0.0000 , 0.0000 , 0.0000 , 0.0023 , 0.0000 , 0.0395 , 0.0000)   ,
{2226}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.3781 , 0.0000 , 0.4237 , 0.0000 , 0.0000 , 0.0000 , 0.0277 , 0.0000 , 0.1288 , 0.0000 , 0.0000 , 0.0000 , 0.0023 , 0.0000 , 0.0395 , 0.0000))) ,
{2311}  ((( 0.0000 , 0.0000 , 0.0000 , 0.0090 , 0.0000 , 0.1102 , 0.0000 , 0.0000 , 0.0000 , 0.0236 , 0.0000 , 0.2297 , 0.0000 , 0.0538 , 0.0000 , 0.0631 , 0.0000 , 0.3008 , 0.0000 , 0.2099)   ,
{2312}    ( 0.0000 , 0.0000 , 0.0010 , 0.0048 , 0.0000 , 0.1245 , 0.0000 , 0.0000 , 0.0148 , 0.0376 , 0.0043 , 0.1906 , 0.0000 , 0.0510 , 0.0273 , 0.0512 , 0.0256 , 0.1992 , 0.0223 , 0.2459)   ,
{2313}    ( 0.0000 , 0.0000 , 0.0062 , 0.0273 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0502 , 0.1056 , 0.0062 , 0.0290 , 0.0000 , 0.0000 , 0.1303 , 0.2931 , 0.0704 , 0.1382 , 0.0070 , 0.1364)   ,
{2314}    ( 0.0000 , 0.0000 , 0.0016 , 0.0791 , 0.0000 , 0.0032 , 0.0000 , 0.0000 , 0.0728 , 0.1297 , 0.0285 , 0.0475 , 0.0000 , 0.0016 , 0.0316 , 0.0665 , 0.0443 , 0.2168 , 0.0364 , 0.2405)   ,
{2315}    ( 0.0000 , 0.0000 , 0.0017 , 0.0380 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0155 , 0.0484 , 0.0225 , 0.1693 , 0.0000 , 0.0069 , 0.0190 , 0.0518 , 0.0656 , 0.2971 , 0.0225 , 0.2418)   ,
{2316}    ( 0.0000 , 0.0000 , 0.0000 , 0.0090 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0151 , 0.0904 , 0.0392 , 0.1777 , 0.0000 , 0.0030 , 0.0120 , 0.0572 , 0.0813 , 0.3012 , 0.0301 , 0.1837))  ,
{2321}   (( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{2322}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0194 , 0.0227 , 0.0002 , 0.0257 , 0.0000 , 0.0000 , 0.0319 , 0.0525 , 0.0212 , 0.1707 , 0.0000 , 0.0000 , 0.0484 , 0.0764 , 0.0823 , 0.4486)   ,
{2323}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0221 , 0.1127 , 0.0008 , 0.0092 , 0.0000 , 0.0000 , 0.0586 , 0.1778 , 0.0066 , 0.0927 , 0.0000 , 0.0000 , 0.0906 , 0.1819 , 0.0329 , 0.2141)   ,
{2324}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0144 , 0.1713 , 0.0000 , 0.0099 , 0.0000 , 0.0000 , 0.0504 , 0.1623 , 0.0108 , 0.1187 , 0.0000 , 0.0000 , 0.0289 , 0.0952 , 0.0470 , 0.2911)   ,
{2325}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0067 , 0.0976 , 0.0000 , 0.0202 , 0.0000 , 0.0000 , 0.0337 , 0.0842 , 0.0168 , 0.1481 , 0.0000 , 0.0000 , 0.0471 , 0.0640 , 0.0774 , 0.4040)   ,
{2326}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0067 , 0.0976 , 0.0000 , 0.0202 , 0.0000 , 0.0000 , 0.0337 , 0.0842 , 0.0168 , 0.1481 , 0.0000 , 0.0000 , 0.0471 , 0.0640 , 0.0774 , 0.4040))) ,
{2411}  ((( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0034 , 0.0000 , 0.0328 , 0.0000 , 0.0844 , 0.0000 , 0.0102 , 0.0000 , 0.2925 , 0.0000 , 0.5767)   ,
{2412}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0026 , 0.0065 , 0.0309 , 0.0578 , 0.0000 , 0.0938 , 0.0045 , 0.0139 , 0.0345 , 0.2347 , 0.0155 , 0.5053)   ,
{2413}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0124 , 0.0095 , 0.0819 , 0.0855 , 0.0000 , 0.0021 , 0.0034 , 0.0250 , 0.0781 , 0.6067 , 0.0096 , 0.0858)   ,
{2414}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0026 , 0.0216 , 0.0104 , 0.2197 , 0.0000 , 0.0182 , 0.0143 , 0.0489 , 0.0229 , 0.3348 , 0.0411 , 0.2656)   ,
{2415}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0123 , 0.0775 , 0.0034 , 0.0953 , 0.0000 , 0.0041 , 0.0158 , 0.0686 , 0.0556 , 0.2428 , 0.0912 , 0.3333)   ,
{2416}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0096 , 0.0795 , 0.0193 , 0.0169 , 0.0000 , 0.0024 , 0.0289 , 0.0217 , 0.0410 , 0.2169 , 0.0747 , 0.4892))  ,
{2421}   (( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{2422}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0198 , 0.0134 , 0.0854 , 0.0996 , 0.0000 , 0.0000 , 0.0267 , 0.0510 , 0.0910 , 0.6130)   ,
{2423}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0245 , 0.0235 , 0.0591 , 0.1254 , 0.0000 , 0.0000 , 0.0282 , 0.1355 , 0.0660 , 0.5378)   ,
{2424}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0079 , 0.0471 , 0.0116 , 0.2487 , 0.0000 , 0.0000 , 0.0063 , 0.0880 , 0.0379 , 0.5525)   ,
{2425}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0080 , 0.1072 , 0.0067 , 0.1984 , 0.0000 , 0.0000 , 0.0214 , 0.0858 , 0.1019 , 0.4705)   ,
{2426}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0080 , 0.1072 , 0.0067 , 0.1984 , 0.0000 , 0.0000 , 0.0214 , 0.0858 , 0.1019 , 0.4705)))),
{3111} (((( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{3112}    ( 0.0195 , 0.0000 , 0.0420 , 0.0000 , 0.1261 , 0.0000 , 0.0000 , 0.0000 , 0.0105 , 0.0000 , 0.2643 , 0.0000 , 0.2222 , 0.0000 , 0.0045 , 0.0000 , 0.0661 , 0.0000 , 0.2447 , 0.0000)   ,
{3113}    ( 0.2514 , 0.0000 , 0.1886 , 0.0000 , 0.1886 , 0.0000 , 0.0000 , 0.0000 , 0.0229 , 0.0000 , 0.1029 , 0.0000 , 0.0971 , 0.0000 , 0.0114 , 0.0000 , 0.0686 , 0.0000 , 0.0686 , 0.0000)   ,
{3114}    ( 0.2785 , 0.0000 , 0.2034 , 0.0000 , 0.1186 , 0.0000 , 0.0000 , 0.0000 , 0.0387 , 0.0000 , 0.1186 , 0.0000 , 0.0993 , 0.0000 , 0.0048 , 0.0000 , 0.0242 , 0.0000 , 0.1138 , 0.0000)   ,
{3115}    ( 0.3738 , 0.0000 , 0.2025 , 0.0000 , 0.1236 , 0.0000 , 0.0000 , 0.0000 , 0.0498 , 0.0000 , 0.1329 , 0.0000 , 0.0571 , 0.0000 , 0.0021 , 0.0000 , 0.0083 , 0.0000 , 0.0498 , 0.0000)   ,
{3116}    ( 0.4764 , 0.0000 , 0.0849 , 0.0000 , 0.1309 , 0.0000 , 0.0000 , 0.0000 , 0.0106 , 0.0000 , 0.1132 , 0.0000 , 0.0802 , 0.0000 , 0.0035 , 0.0000 , 0.0342 , 0.0000 , 0.0660 , 0.0000))  ,
{3121}   (( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{3122}    ( 0.0000 , 0.0942 , 0.0000 , 0.0000 , 0.0639 , 0.0000 , 0.2266 , 0.0000 , 0.0000 , 0.0000 , 0.0140 , 0.0000 , 0.3462 , 0.0000 , 0.0000 , 0.0000 , 0.0023 , 0.0000 , 0.2529 , 0.0000)   ,
{3123}    ( 0.0000 , 0.3983 , 0.0000 , 0.0000 , 0.1607 , 0.0000 , 0.2531 , 0.0000 , 0.0000 , 0.0000 , 0.0171 , 0.0000 , 0.1173 , 0.0000 , 0.0000 , 0.0000 , 0.0026 , 0.0000 , 0.0510 , 0.0000)   ,
{3124}    ( 0.0000 , 0.4125 , 0.0000 , 0.0000 , 0.1821 , 0.0000 , 0.1551 , 0.0000 , 0.0000 , 0.0000 , 0.0189 , 0.0000 , 0.1570 , 0.0000 , 0.0000 , 0.0000 , 0.0048 , 0.0000 , 0.0697 , 0.0000)   ,
{3125}    ( 0.0000 , 0.5523 , 0.0000 , 0.0000 , 0.1526 , 0.0000 , 0.1057 , 0.0000 , 0.0000 , 0.0000 , 0.0238 , 0.0000 , 0.1265 , 0.0000 , 0.0000 , 0.0000 , 0.0018 , 0.0000 , 0.0374 , 0.0000)   ,
{3126}    ( 0.0000 , 0.5523 , 0.0000 , 0.0000 , 0.1526 , 0.0000 , 0.1057 , 0.0000 , 0.0000 , 0.0000 , 0.0238 , 0.0000 , 0.1265 , 0.0000 , 0.0000 , 0.0000 , 0.0018 , 0.0000 , 0.0374 , 0.0000))) ,
{3211}  ((( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{3212}    ( 0.0000 , 0.0000 , 0.0043 , 0.0000 , 0.0429 , 0.0000 , 0.0000 , 0.0000 , 0.0206 , 0.0000 , 0.1309 , 0.0000 , 0.3329 , 0.0000 , 0.0062 , 0.0000 , 0.0570 , 0.0000 , 0.4052 , 0.0000)   ,
{3213}    ( 0.0000 , 0.0000 , 0.0683 , 0.0000 , 0.5524 , 0.0000 , 0.0000 , 0.0000 , 0.0254 , 0.0000 , 0.1508 , 0.0000 , 0.0905 , 0.0000 , 0.0063 , 0.0000 , 0.0365 , 0.0000 , 0.0698 , 0.0000)   ,
{3214}    ( 0.0000 , 0.0000 , 0.1420 , 0.0000 , 0.4881 , 0.0000 , 0.0000 , 0.0000 , 0.0244 , 0.0000 , 0.1046 , 0.0000 , 0.1344 , 0.0000 , 0.0032 , 0.0000 , 0.0187 , 0.0000 , 0.0846 , 0.0000)   ,
{3215}    ( 0.0000 , 0.0000 , 0.4649 , 0.0000 , 0.3663 , 0.0000 , 0.0000 , 0.0000 , 0.0287 , 0.0000 , 0.0576 , 0.0000 , 0.0413 , 0.0000 , 0.0010 , 0.0000 , 0.0084 , 0.0000 , 0.0318 , 0.0000)   ,
{3216}    ( 0.0000 , 0.0000 , 0.6418 , 0.0000 , 0.0692 , 0.0000 , 0.0000 , 0.0000 , 0.0341 , 0.0000 , 0.0711 , 0.0000 , 0.0939 , 0.0000 , 0.0011 , 0.0000 , 0.0090 , 0.0000 , 0.0797 , 0.0000))  ,
{3221}   (( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{3222}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0119 , 0.0000 , 0.2330 , 0.0000 , 0.0000 , 0.0000 , 0.0116 , 0.0000 , 0.4006 , 0.0000 , 0.0000 , 0.0000 , 0.0021 , 0.0000 , 0.3408 , 0.0000)   ,
{3223}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0497 , 0.0000 , 0.7423 , 0.0000 , 0.0000 , 0.0000 , 0.0122 , 0.0000 , 0.1304 , 0.0000 , 0.0000 , 0.0000 , 0.0015 , 0.0000 , 0.0640 , 0.0000)   ,
{3224}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0959 , 0.0000 , 0.5600 , 0.0000 , 0.0000 , 0.0000 , 0.0162 , 0.0000 , 0.2235 , 0.0000 , 0.0000 , 0.0000 , 0.0025 , 0.0000 , 0.1019 , 0.0000)   ,
{3225}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.2745 , 0.0000 , 0.5046 , 0.0000 , 0.0000 , 0.0000 , 0.0167 , 0.0000 , 0.1436 , 0.0000 , 0.0000 , 0.0000 , 0.0020 , 0.0000 , 0.0586 , 0.0000)   ,
{3226}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.2745 , 0.0000 , 0.5046 , 0.0000 , 0.0000 , 0.0000 , 0.0167 , 0.0000 , 0.1436 , 0.0000 , 0.0000 , 0.0000 , 0.0020 , 0.0000 , 0.0586 , 0.0000))) ,
{3311}  ((( 0.0000 , 0.0000 , 0.0000 , 0.0018 , 0.0000 , 0.0792 , 0.0000 , 0.0000 , 0.0000 , 0.0155 , 0.0000 , 0.1573 , 0.0000 , 0.0236 , 0.0000 , 0.0995 , 0.0000 , 0.3801 , 0.0000 , 0.2430)   ,
{3312}    ( 0.0000 , 0.0000 , 0.0000 , 0.0063 , 0.0000 , 0.0888 , 0.0000 , 0.0000 , 0.0055 , 0.0173 , 0.0165 , 0.1909 , 0.0000 , 0.0353 , 0.0259 , 0.0487 , 0.0369 , 0.2372 , 0.0228 , 0.2679)   ,
{3313}    ( 0.0000 , 0.0000 , 0.0000 , 0.0064 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.1674 , 0.0622 , 0.0236 , 0.0129 , 0.0000 , 0.0064 , 0.1974 , 0.3648 , 0.0064 , 0.0579 , 0.0129 , 0.0815)   ,
{3314}    ( 0.0000 , 0.0000 , 0.0216 , 0.0378 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0811 , 0.1622 , 0.0054 , 0.0919 , 0.0000 , 0.0054 , 0.0973 , 0.1946 , 0.0324 , 0.0865 , 0.0162 , 0.1676)   ,
{3315}    ( 0.0000 , 0.0000 , 0.0000 , 0.0272 , 0.0000 , 0.0068 , 0.0000 , 0.0000 , 0.0204 , 0.0068 , 0.0136 , 0.0680 , 0.0000 , 0.0000 , 0.0272 , 0.0680 , 0.0408 , 0.3810 , 0.0408 , 0.2993)   ,
{3316}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0097 , 0.0000 , 0.0777 , 0.0000 , 0.0000 , 0.0000 , 0.0680 , 0.0777 , 0.3010 , 0.0194 , 0.4466))  ,
{3321}   (( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{3322}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0042 , 0.0006 , 0.0000 , 0.0158 , 0.0000 , 0.0000 , 0.0291 , 0.0352 , 0.0170 , 0.0873 , 0.0000 , 0.0000 , 0.0570 , 0.0485 , 0.0904 , 0.6149)   ,
{3323}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0121 , 0.0452 , 0.0000 , 0.0016 , 0.0000 , 0.0000 , 0.0987 , 0.1000 , 0.0062 , 0.0524 , 0.0000 , 0.0000 , 0.1849 , 0.2386 , 0.0275 , 0.2327)   ,
{3324}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0077 , 0.1386 , 0.0000 , 0.0095 , 0.0000 , 0.0000 , 0.0922 , 0.1261 , 0.0018 , 0.0547 , 0.0000 , 0.0000 , 0.0607 , 0.1190 , 0.0416 , 0.3480)   ,
{3325}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0385 , 0.1374 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0275 , 0.0549 , 0.0165 , 0.0604 , 0.0000 , 0.0000 , 0.0055 , 0.0330 , 0.0714 , 0.5549)   ,
{3326}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0385 , 0.1374 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0275 , 0.0549 , 0.0165 , 0.0604 , 0.0000 , 0.0000 , 0.0055 , 0.0330 , 0.0714 , 0.5549))) ,
{3411}  ((( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0016 , 0.0000 , 0.0237 , 0.0000 , 0.0870 , 0.0000 , 0.0071 , 0.0000 , 0.2613 , 0.0000 , 0.6193)   ,
{3412}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0022 , 0.0036 , 0.0139 , 0.0385 , 0.0000 , 0.1184 , 0.0000 , 0.0109 , 0.0130 , 0.1738 , 0.0064 , 0.6193)   ,
{3413}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0075 , 0.0045 , 0.0831 , 0.0748 , 0.0000 , 0.0002 , 0.0016 , 0.0214 , 0.1088 , 0.6221 , 0.0037 , 0.0723)   ,
{3414}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0052 , 0.0342 , 0.0312 , 0.2077 , 0.0000 , 0.0208 , 0.0004 , 0.0558 , 0.0108 , 0.4392 , 0.0095 , 0.1852)   ,
{3415}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0016 , 0.0643 , 0.0033 , 0.0791 , 0.0000 , 0.0165 , 0.0033 , 0.0643 , 0.0231 , 0.2026 , 0.1087 , 0.4333)   ,
{3416}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0070 , 0.0000 , 0.0035 , 0.0000 , 0.0000 , 0.0000 , 0.0458 , 0.0176 , 0.2606 , 0.0634 , 0.6021))  ,
{3421}   (( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000)   ,
{3422}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0060 , 0.0022 , 0.0582 , 0.0921 , 0.0000 , 0.0000 , 0.0068 , 0.0114 , 0.0686 , 0.7548)   ,
{3423}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0198 , 0.0126 , 0.1034 , 0.1159 , 0.0000 , 0.0000 , 0.0289 , 0.1114 , 0.0866 , 0.5215)   ,
{3424}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0084 , 0.0352 , 0.0155 , 0.2433 , 0.0000 , 0.0000 , 0.0048 , 0.0816 , 0.0220 , 0.5892)   ,
{3425}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0083 , 0.0812 , 0.0141 , 0.2565 , 0.0000 , 0.0000 , 0.0035 , 0.0435 , 0.0518 , 0.5412)   ,
{3426}    ( 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0000 , 0.0083 , 0.0812 , 0.0141 , 0.2565 , 0.0000 , 0.0000 , 0.0035 , 0.0435 , 0.0518 , 0.5412)))));

pAgeCorr:array[1..NumberOfAgeGroups] of integer=(1,1,1, 2,2,2, 3,3,3, 4,4,4, 5,5,5, 6,6);

var hhds:array[1..NumberOfAreaTypes,1..nHInc,1..nHComp,0..NumberOfTimeSteps] of single;
demVar,Subregion,AreaType,HInc,HComp,ts,WorkerGr,IncomeGr,EthnicGr,HhldType,AgeGroup:integer;  ouf:text;

begin

   for AreaType:=1 to NumberOfAreaTypes do
   for HInc:=1 to nHInc do
   for HComp:=1 to nHComp do
   for ts:=0 to NumberOfTimeSteps do
    hhds[AreaType,HInc,HComp,ts]:=0;

   for ts:=0 to NumberOfTimeSteps do begin
      for AreaType:=1 to NumberOfAreaTypes do
      for WorkerGr:=1 to NumberOfWorkerGrs do
      for IncomeGr:=1 to NumberOfIncomeGrs do
      for EthnicGr:=1 to NumberOfEthnicGrs do
      for HhldType:=1 to NumberOfHhldTypes do
      for AgeGroup:=1 to NumberOfAgeGroups do begin
{worker code is backwards - switch below}
        HInc:=IncomeGr;
        for HComp:=1 to nHComp do begin
          hhds[AreaType,HInc,HComp,ts]:=
          hhds[AreaType,HInc,HComp,ts]+
           Population[AreaType][AgeGroup][HhldType][EthnicGr][IncomeGr][WorkerGr][ts]
         * phFrac[IncomeGr,HhldType,3-WorkerGr,pAgeCorr[AgeGroup],HComp]
         / (AvgHSize[Hinc,HComp]*HSizeAdj[HCompSize[HComp]]);
        end;
     end;
   end;

   assign(ouf,OutputDirectory+RunLabel+'hhlds.csv'); rewrite(ouf);


   write(ouf,'Year');
   for ts:=0 to NumberOfTimeSteps do if (ts*TimeStepLength=round(ts*TimeStepLength)) then begin
     write(ouf,',',StartYear + ts*TimeStepLength:4:0);
   end;
   writeln(ouf);

   for AreaType:=1 to NumberOfAreaTypes do
   for HInc:=1 to nHInc do
   for HComp:=1 to nHComp do begin
     write(ouf,AreaTypeLabels[AreaType],'/',
     HIncLabel[HInc],'/',
     HSizeLabel[HCompSize[HComp]],'/',
     HWorkLabel[HCompWork[HComp]+1],'/',
     HKidsLabel[HCompKids[HComp]+1]);

     for ts:=0 to NumberOfTimeSteps do if (ts*TimeStepLength=round(ts*TimeStepLength)) then
       write(ouf,',',hhds[AreaType,HInc,HComp,ts]:1:0);
     writeln(ouf);
   end;
   writeln(ouf);

{write some outputs for calibration}
   demVar := 0;  {population by subregion}
   ts:=round(5.0/TimeStepLength);
   CalculateDemographicMarginals(demVar,ts);

   write(ouf,'2015 Output by Age Group');
   for subregion:=1 to NumberOfSubregions do write(ouf,',',SubregionLabels[subregion]);
   writeln(ouf);
   for AgeGroup:=1 to NumberOfAgeGroups do begin
     write(ouf,AgeGroupLabels[AgeGroup]);
     for subregion:=1 to NumberOfSubregions do write(ouf,',',AgeGroupMarginals[subregion][AgeGroup][ts]:1:0);
     writeln(ouf);
   end;
   writeln(ouf); writeln(ouf);

   write(ouf,'2015 Output by HHld Type');
   for subregion:=1 to NumberOfSubregions do write(ouf,',',SubregionLabels[subregion]);
   writeln(ouf);
   for HhldType:=1 to NumberOfHhldTypes do begin
     write(ouf,HhldTypeLabels[HhldType]);
     for subregion:=1 to NumberOfSubregions do write(ouf,',',HhldTypeMarginals[subregion][HhldType][ts]:1:0);
     writeln(ouf);
   end;
   writeln(ouf); writeln(ouf);

   write(ouf,'2015 Output by Ethnic Group');
   for subregion:=1 to NumberOfSubregions do write(ouf,',',SubregionLabels[subregion]);
   writeln(ouf);
   for EthnicGr:=1 to NumberOfEthnicGrs do begin
     write(ouf,EthnicGrLabels[EthnicGr]);
     for subregion:=1 to NumberOfSubregions do write(ouf,',',EthnicGrMarginals[subregion][EthnicGr][ts]:1:0);
     writeln(ouf);
   end;
   writeln(ouf); writeln(ouf);

   write(ouf,'2015 Output by Income Group');
   for subregion:=1 to NumberOfSubregions do write(ouf,',',SubregionLabels[subregion]);
   writeln(ouf);
   for IncomeGr:=1 to NumberOfIncomeGrs do begin
     write(ouf,IncomeGrLabels[IncomeGr]);
     for subregion:=1 to NumberOfSubregions do write(ouf,',',IncomeGrMarginals[subregion][IncomeGr][ts]:1:0);
     writeln(ouf);
   end;
   writeln(ouf); writeln(ouf);

   write(ouf,'2015 Output by Workforce Part');
   for subregion:=1 to NumberOfSubregions do write(ouf,',',SubregionLabels[subregion]);
   writeln(ouf);
   for WorkerGr:=1 to NumberOfWorkerGrs do begin
     write(ouf,WorkerGrLabels[WorkerGr]);
     for subregion:=1 to NumberOfSubregions do write(ouf,',',WorkerGrMarginals[subregion][WorkerGr][ts]:1:0);
     writeln(ouf);
   end;
   writeln(ouf); writeln(ouf);

   write(ouf,'2015 Output by Area Type');
   for subregion:=1 to NumberOfSubregions do write(ouf,',',SubregionLabels[subregion]);
   writeln(ouf);
   for AreaType:=1 to NumberOfAreaTypes do begin
     write(ouf,AreaTypeLabels[AreaType]);
     for subregion:=1 to NumberOfSubregions do write(ouf,',',AreaTypeMarginals[subregion][AreaType][ts]:1:0);
     writeln(ouf);
   end;
   writeln(ouf); writeln(ouf);

   close(ouf);

end;

{Main simulation program}
var demVar:integer;

begin
 {Read in all input data}
  writeln('Loading input data from spreadsheet');
  ReadUserInputData;

  if testWriteYear>0 then begin
    assign(outest,'test_out.csv'); rewrite(outest);
    writeln(outest,'Year,AreaType,AgeGroup,HhldType,EthnicGr,IncomeGr,WorkerGr,PopulationNew,PopulationOld,',
     'AgeingOut,DeathsOut,MarriagesOut,DivorcesOut,FirstChildOut,',
     'EmptyNestOut,LeaveNestOut,AcculturationOut,WorkerStatusOut,IncomeGroupOut,',
     'ForeignOutMigration,DomesticOutMigration,RegionalOutMigration,',
     'AgeingIn,BirthsIn,MarriagesIn,DivorcesIn,FirstChildIn,',
     'EmptyNestIn,LeaveNestIn,AcculturationIn,WorkerStatusIn,IncomeGroupIn,',
     'ForeignInMigration,DomesticInMigration,RegionalInMigration,',
     'ExogPopChange,HUrbEnter,HUrbLeave,IUrbEnter,IUrbLeave');
   end;

{Initialize all sectors}
  writeln('Using IPF for base year population');
  InitializePopulation;
  {InitializeEmployment;  not necessary - in input data }
  {InitializeLandUse:     not necessary - in input data }
  {InitializeTransportationSupply;  not necessary - in input data }

  {Do travel demand for year 0 without supply feedback effects}
    CalculateTravelDemand(0);

  {step through time t}
  write('Simulating year ... ');
  TimeStep:=0;
  for demVar:=1 to NumberOfDemographicVariables do CalculateDemographicMarginals(demVar,TimeStep);
  repeat
    TimeStep := TimeStep + 1;
    Year := Year + TimeStepLength;
    if Year = trunc(Year) then write(Year:8:0);

    {Calculate feedbacks between sectors based on levels from previous time step}
    CalculateDemographicFeedbacks(TimeStep);
    CalculateEmploymentFeedbacks(TimeStep);
    CalculateLandUseFeedbacks(TimeStep);
    CalculateTransportationSupplyFeedbacks(TimeStep);

    {Calculate rate variables for time t based on levels from time t-1 and feedback effects}
    CalculateDemographicTransitionRates(TimeStep);
    CalculateEmploymentTransitionRates(TimeStep);
    CalculateLandUseTransitionRates(TimeStep);
    CalculateTransportationSupplyTransitionRates(TimeStep);

    {Apply transition rates to get new levels for time t}
    ApplyDemographicTransitionRates(TimeStep);
    ApplyEmploymentTransitionRates(TimeStep);
    ApplyLandUseTransitionRates(TimeStep);
    ApplyTransportationSupplyTransitionRates(TimeStep);

    {based on resulting population for time t, calculate the travel demand}
    CalculateTravelDemand(TimeStep);

    for demVar:=1 to NumberOfDemographicVariables do CalculateDemographicMarginals(demVar,TimeStep);

  until TimeStep >= NumberOfTimeSteps; {end of simulation}

  {Write out all simulation results}
  writeln;
  writeln('Writing results files ....');

  WriteSimulationResults;
  WriteHouseholdConversionFile;

  {TimeStep:=0; repeat TimeStep:=TimeStep+1 until TimeStep = 9999999;}
  if testWriteYear>0 then close(outest);
  {write('Simulation finished. Press Enter to send results to Excel'); readln;}

end.



