# Garmin-Connect-Workout-and-Schedule-creator
Create Garmin Connect workouts with a "turbo language" and schedule them.

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

The Workout creator uses a Excel spreadsheet for defining scheduling workouts, describe workouts and workouttypes. The Excel spreadsheet has 3 obligatory worksheets called "Schedule", "Workout" and "workoutType", don't change the names of these 3 worksheets !

The worksheet 
