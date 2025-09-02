# SHGScan
SharpCap Python script to automatically scan the sun, for use with spectroheliographs

This IronPython script is intended for use with SharpCap v4.0 and higher

The intended function of this script is to automate image acquisition of the sun with a spectroheliograph such as the MLAstro SHG-700, in conjunction with an ASCOM compliant mount and a high frame rate camera, all controlled by SharpCap. It automatically detects the edges of the sun and calculates the correct slew speed based on the preview frame rate to achieve a close to 1.0 Y:X ratio for reconstruction
This script installs a custom button on the SharpCap toolbar labeled "|   SHG Scan   |"
When clicked, this displays a Form with a number of controls on it. The Form runs asynchronously, so you can continue to interact with SharpCap UI elements while it is running. The intended workflow is:
  Start SharpCap, connect to imaging camera, locate sun with mount, set appropriate spectroheliograph imaging ROI. This is typically wide but not very tall, in order to maximize frame rate. It must be at least 100 pixels high for the script.
  Camera should be set to MONO16, .SER format. The script will enforce this. May provide an option to set different values at some point if there is a convincing reason to do so.
  Set appropriate exposure and gain to maximize historgram fill without clipping.
  On startup, the form retrieves the frame rate from SharpCap - the status message currently must display a frame rate, e.g. "Previewing: xxx frames at xxx fps". If the frame rate is too slow, it will only show the number of frames, however this is likely to far too slow for spectroheliograph use in any case. 
  The form then calculates the expected slew speed and acquisition time, based on default sun width estimate of 2300 pixels (roughly correct for 80mm f/7 scope, ZWO ASI678MM (2 micron pixels), and SHG-700)
  Manually scan the mount to center the widest part of the imaged spectrum
  Click "Measure Sun" to measure correct sun width and refine slew speed calculation
  Set parameters as desired:
    1. Slew RA - default is to slew the mount in RA for data acquisition. Bump slews will slew in Dec. Unchecking this box reverses these preferences
    2. Number of cycles - how many passes to make over the sun (forward and return)
    3. Bidirectional - if checked, will acquire data in both directions of slew. If unchecked, will return the mount at high speed (8x the forward scan speed)
    4. Slew pad - how many seconds to slew past the edge transition (both leading and trailing edges)
    5. Cycle sleep - how many seconds to sleep between scan cycles
    6. Bump rate - rate for corrective 0.25s / 0.5s slews to compensate for solar drift, performed between cycles
    7. Measure Sun - shows current spectrum image width used for calculation. Can be manually overridden if desired
    8. decenter - read only, shows offset of spectrum from center
    9. Frame rate - capture camera frame rate. Can be manually overridden if desired
  Click "Go"
  The script will now slew the mount to the starting position, detecting the light falloff passing the solar limb to the start position, then for an additional Slew Pad seconds. This pad is useful to allow the mount axis to come up to speed smoothly, and to avoid clipping prominences and other dim features past the limb. It does this by examining the average value in a 100x100 pixel area in the center of the ROI and waiting for it to fall below 10% brightness
  The script then starts image capture, mount then slews forward at the calculated rate, again looking for a bright to dark transition at the trailing edge, then for an additional Slew Pad seconds. Image capture is then stopped. Note that SharpCap will block until output file write is complete
  If Bidirectional is checked, the script will restart image capture, then rescan in the opposite direction to the starting position at the same rate
  If Bidirectional is not checked, the script will reposition the mount at high speed (8x the forward scan rate) until it passes the solar limb, then for an additional Slew Pad seconds at the forward scan rate.
  Monitor spectrum to ensure that it remains centered. It may be helpful to display the target reticle on the image to ensure that the disappearing vertical spectrum stripe remains centered.
  If the spectrum appears to drift left or right, click the "<-" or "->" buttons to request a 1/4 sec corrective slew between cycles. "<<" and ">>" perform 1/2 second slews.
  If the spectrum moves in the opposite direction from desired, clicking the "Swap" button will reverse the bump slew directions.
  When all cycles are complete, the script returns the mount to approximately the solar midpoint, by reversing the slew for half the time measured to accomplish the most recent forward scan.
  If the "Abort" button is clicked, acquisition is terminated and the mount is repositioned to the starting position prior to _any_ scan cycles. Note that if alignment is not precise, this may not be at the midpoint of sun any more

Currently all captured sequences are stored in the default SharpCap capture folder. It may be helpful to rename this folder after each group of cycles in order to avoid confusion as to which files belong to which run. At some future point, the script may provide some means to specify a folder name stub, and automatically organize the runs into sub folders.

Patrick Hsieh
flankeronetwo@gmail.com
