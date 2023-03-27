# Garmin-Connect-Workout-and-Schedule-creator
Create Garmin Connect workouts with a "turbo language" and schedule them.

**Language**
The Workouts creator uses the following "language" to define a workout
  0-9  defines distance or time, f.e. 100m or 15:30 minutes
  m    indicator that de numbers before the "m" is a distance, f.e. 100m
  :    indicator that de 2 numbers before and after the ":" are a time-definition, f.e. 15:30
  Z    zone, f.e. Z0 for Zone zero
  *    indicator that the number after the "*" indicates the number of iterations of the repeat group
  !    indicator that the digits after the "!" indicate the Stroketype or Steptype, f.e. !SS is stroketype "Schoolslag" (dutch for breaststroke) 
       or !WUP for steptype 'Warming up'
  @    indicator that the digits after the "@" indicate a Zone, RPM or BPM, f.e. 100bpm or 80rpm
  (    start of a repeat group, f.e. (100m!SS + 00:20!RUST) * 2 (RUST is dutch for Rest)
  )    end of a repeat group
  +    next workoutstep
  
Couple of examples:
  Swimming: 300m!ES@Z1 + (200m!BC@Z1 + 50m!SS + 00:20!RUST)*4 + (25m!BC@Z3)*6 + 250m!SS
  Running:  15:00!WUP + (05:00@Z2 + 03:00!RUST) * 4 + 10:00@Z1 + (01:00@Z3 + 01:00!RUST) * 3 + 15:00!COOL
  Cycling:  90:00@Z1@100rpm

**Excel sheet**
The Workout creator uses a Excel spreadsheet for defining scheduling workouts, describe workouts and workouttypes. The Excel spreadsheet has 3 obligatory worksheets called "Schedule", "Workout" and "workoutType", don't change the names of these 3 worksheets !

The worksheet "Schedule" contains the following columns:
  Date:         Date for scheduling the workout in your Garmin Connect Calendar
  Workout:      Lookup field for a defined workout in the Workout worksheet
  Comments:     Commenting your Workout, this comment is also visible in your Garmin Connect Workout. Max. 512 characters
  
The worksheet "Workout" contains the following columns:
  Name:         Name of the workout, max. 80 characters
  sportType:    Type of sport. Import to add also this sportType to the worksheet workoutType (explained hereunder)
  Description:  The workout description based on the workout creator language as explained here above
  
The worksheet "workoutType" contains the following columns:
  Name:         Short name of the workout type
  Value:        Reference to the Garmin Connect API definitions f.e. swimStrokeType, stepTypeId of sportTypeId. Access to the Garmin API documentation is only
		            granted to professional developers, so getting these values is only possible via reversed engineering. So define a workout with the appropiate
		            sport type, step type and/of swimstroke type. Export this workout to json format and open the downloaded file in a editor/viewer. Determine the
		            appropiate values for sportTypeId, stepTypeId or swimStrokeTypeId or other Id's and coupled workoutType
  apiType:		  definition of the Id type, f.e. sportTypeId or stepTypeId etc.
  workoutType:	definition of the workoutType, also use reversed engineering to find out the appropiate values
  Description:	Free description, is not used by the workout creator

**Reversed engeneering a new workout type**
If you want to add a new Sporttype, for example "Yoga", you can follow the following steps.
1)  Create an Yoga workout in Garmin Connect, filled with your favorite trainingtypes as workout steps. Save it.
2)  Download the workout to your desktop and open it with notepad
3)  Select all text and copy it to your clipboard
4)  Open https://jsonlint.com/ and paste the text, click on "Validate JSON". JSONLint tides and validates the messy JSON code.
5)  Search for sportTypeId. In this example the sportTypeId is "7" and the sportTypeKey = "yoga"
6)  

**Known issues**
Multisport workouts, like Duathlon of Triathlon ain't supported in Garmin Connect, add seperate workouts as work-arround
