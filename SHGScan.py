# *****************************************************************************************************
#     Automated scanning / capture for spectroheliographs
# (c) 2025 Patrick Hsieh
#
# Version: 1.0 (9/1/2025) Initial release
#
# To Do: 
    # Organize captures into folders
    # Automatically scan for widest point
    # ? C++ routine for faster width / edge measurement
    # ? Other edge detectio algorithms - horizontal slice might be more robust against decentering
    # ? Automated sun finding? should be possible to slew back and forth to maximize the ROI mean brightness, as long as somewhat close to the sun
    # ? save start position at start of each cycle? On abort reposition here, then slew half of last cycle time to reach midpoint. If no cycles yet completed
    #    would simply reposition to start coordinates.
# *****************************************************************************************************

import time, os, sys, math, clr, io, re
from pathlib import Path
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Threading.Tasks")
from System.Threading.Tasks import Task
import System.Drawing
import System.Windows.Forms

from System import EventHandler
from System.Drawing import *
from System.Drawing.Drawing2D import InterpolationMode
from System.Windows.Forms import *
from SharpCap.Base import Interfaces, NotificationStatus

############### GLOBALS #################
MAX_BRIGHT = 65535
DEFAULT_SUN_WIDTH = 2300        # roughly correct for 80mm f/7 refractor, ASI678MM (2u pixels)
DEFAULT_CYCLE_SLEEP=0.5
DEFAULT_SLEWPAD = 0.5
DEFAULT_THRESHOLD = 0.1
DEFAULT_CYCLES = 15
DEFAULT_BIDIRECTIONAL = False
DEFAULT_BUMPSWAP = False
DEFAULT_AXISTOMOVE = 0          # RA by default
IsAbortState = False
MainForm = None
globalTabIndex=0
### END GOBALS


class SHGForm(Form):

    # User input vars
    SlewFactor = 1
    SlewPad = DEFAULT_SLEWPAD
    CycleSleep = DEFAULT_CYCLE_SLEEP
    NumCycles = DEFAULT_CYCLES
    Bidirectional = DEFAULT_BIDIRECTIONAL
    BumpRate = 8
    BumpSwap = DEFAULT_BUMPSWAP
    AxisToMove = DEFAULT_AXISTOMOVE      # RA by default. Dec is axis 1
    
    # Vars for frame handling
    EdgePassed = False
    PositiveSignal = False
    LimbThreshold = MAX_BRIGHT * DEFAULT_THRESHOLD
    FrameInterval = 10         # assess for transition every 10th frame
    FrameCount = FrameInterval
    ROIX=0
    ROIY=0
    FrameHandlingDone = False
    SunWidth = 2300
    SunDecenter = 0

    # Bump slew vars
    AmSlewing = False
    BumpSlew = 0
    FrameRate = -1
    
    # Asynchronous flags
    TaskAbortFlag = False

    # SharpCap objects
    SavedCoords = None
    
    def __init__(self):
        self.SuspendLayout()
        self.getSettings();
        self.InitializeComponent()
        self.setupForm()
        self.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        self.AutoScaleDimensions = SizeF(96, 96)
        self.ResumeLayout()
        self.enableGo()

    def InitializeComponent(self):
        self.Text = "Spectroheliograph Auto Scan"
        self.ClientSize = System.Drawing.Size(340, 364)
        self.TopMost = True

    # read settings from file
    def getSettings(self):
        appDir = os.getenv('APPDATA')
        configFn = Path(appDir + "\\SharpCap\\SHG.cfg")
        if (configFn.exists()):     # read in values
            try:
                with configFn.open("r") as f:
                    config = f.read()
                    items = config.split('\n')     # split into lines
                    # parse lines
                    for item in items:
                        m = re.search("([a-zA-Z]+)=([0-9.a-zA-Z]+)", item)
                        if not m:
                            SharpCap.ShowNotification(f"Ignored invalid settings keyword {item}", NotificationStatus.Warning)
                            continue
                        key = m.group(1)
                        value = m.group(2)
                        if key == "NumCycles":
                            self.NumCycles=int(value)
                        elif key == "SunWidth":
                            self.SunWidth = int(value)
                        elif key == "CycleSleep":
                            self.CycleSleep=float(value)
                        elif key == "SlewPad":
                            self.SlewPad = float(value)
                        elif key == "LimbThreshold":
                            DEFAULT_THRESHOLD = float(value)
                            self.LimbThreshold = MAX_BRIGHT * float(value)
                        elif key == "Bidirectional":
                            self.Bidirectional = (value == "True")
                        elif key == "BumpSwap":
                            self.BumpSwap = (value == "True")
                        elif key == "BumpRate":
                            self.BumpRate = int(value)
                        elif key == "AxisToMove":
                            self.AxisToMove = int(value)
                    f.close()
            except:
                SharpCap.ShowNotification("Error reading settings file", NotificationStatus.Error)
                
        else:       # write default values to file
            self.saveSettings()
            
    # save settings to file
    def saveSettings(self):
        appDir = os.getenv('APPDATA')
        configFn = Path(appDir + "\\SharpCap\\SHG.cfg")
        config = f"NumCycles={self.NumCycles}\nSunWidth={self.SunWidth}\nCycleSleep={self.CycleSleep:.2f}\nSlewPad={self.SlewPad:.2f}\nLimbThreshold={DEFAULT_THRESHOLD:.2f}\nBidirectional={self.Bidirectional}\nBumpSwap={self.BumpSwap}\nBumpRate={self.BumpRate}\nAxisToMove={self.AxisToMove}"
        try:
            with configFn.open("w") as f:
                f.write(config)
                f.close()
        except:
            SharpCap.ShowNotification("Error writing settings file", NotificationStatus.Error)

    ######################### item constructors #########################
    def addTextBox(self, name, value, x, y, width, height, handler):
        global globalTabIndex
        newItem = TextBox()
        newItem.AutoSize = True
        newItem.Location = Point(x, y)
        newItem.Name = name
        newItem.Size = Size(width, height)      # 20 default height
        newItem.Text = value
        if handler:
            newItem.Leave += handler
        self.Controls.Add(newItem)
        return newItem
        
    def addComboBox(self, name, x, y, valList, selItem, handler):
        global globalTabIndex
        newItem = ComboBox()
        newItem.Text = name
        newItem.Location = Point(x, y)
        newItem.Size = Size(60,10)
        for item in valList:
            newItem.Items.Add(item)
        newItem.SelectedItem  = selItem
        if handler:
            newItem.SelectedIndexChanged += handler
        newItem.DropDownStyle = ComboBoxStyle.DropDownList
        self.Controls.Add(newItem)
        return newItem

    def addLabel(self, value, x, y):
        global globalTabIndex
        newItem = Label()
        newItem.AutoSize = True
        newItem.Location = Point(x, y)
        newItem.Text = value
        self.Controls.Add(newItem)
        return newItem
        
    def addButton(self, name, func, x, y):
        global globalTabIndex
        newItem = Button()
        newItem.Text = name
        newItem.Location = Point(x, y)
        newItem.Click += func
        newItem.AutoSize = True
        self.Controls.Add(newItem)
        return newItem

    def addCheckbox(self, name, x, y, value, handler):
        global globalTabIndex
        newItem = CheckBox()
        newItem.Text = name
        newItem.AutoSize = True
        newItem.Location = Point(x, y)
        newItem.Checked = value
        if handler:
            newItem.CheckedChanged += handler
        self.Controls.Add(newItem)
        return newItem
        
    def addProgressBar(self, name, x, y, width, height, limit):
        newItem = ProgressBar()
        newItem.Text = name
        newItem.AutoSize = False
        newItem.Location = Point(x, y)
        newItem.Size = Size(width, height)
        newItem.Minimum = 1
        newItem.Maximum = limit
        newItem.Value = 1
        newItem.Step = 1
        newItem.Visible = True
        self.Controls.Add(newItem)
        return newItem

    ######################### input event handlers #########################
    def doNumCyclesChange(self, sender, args):
        try:
            n = int(self.numCycles.Text)
            if (n >=0):
                self.NumCycles = n
                self.progBar.Maximum = n
            else:
                sender.Undo()
        except:
            sender.Undo()

    def doBidirectionalChange(self, sender, args):
        self.Bidirectional = self.bidirectional.Checked

    def doSlewPadChange(self, sender, args):
        try:
            n = float(self.slewPad.Text)
            if (n > 0):
                self.SlewPad = n
            else:
                sender.Undo()
        except:
            sender.Undo()

    def doCycleSleepChange(self, sender, args):
        try:
            n = float(self.cycleSleep.Text)
            if (n > 0):
                self.CycleSleep = n
            else:
                sender.Undo()
        except:
            sender.Undo()

    def doBumpRateChange(self, sender, args):
        self.BumpRate = int(self.bumpRate.SelectedItem.strip(" x"))

    def doBumpSwapChange(self, sender, args):
        self.BumpSwap = self.bumpSwap.Checked

    def doSunWidthChange(self, sender, args):
        try:
            n = int(self.sunWidth.Text)
            if (n > 100):
                self.SunWidth = n
            else:
                sender.Undo()
        except:
            sender.Undo()

    def doFrameRateChange(self, sender, args):
        try:
            n = float(self.frameRate.Text)
            if (n > 0):
                self.FrameRate = n
                self.CalcScanParams()
            else:
                if (self):
                    sender.Undo()
        except:
            if (self):
                sender.Undo()

    # If checked, slew in RA (axis 0), else slew in Dec (axis 1)
    def doAxisToMoveChange(self, sender, args):
        try:
            if (self.axisToMove.Checked):
                self.AxisToMove = 0
            else:
                self.AxisToMove = 1
        except:
            pass
            
    ######################### SHGForm action handlers #########################
    # find the frame rate
    # The method of accumulating frames for 1 second doesn't seem to consistently  produce accurate results, not sure if this is because it takes some
    # time for LatestStatus to be updated or some other reason. Robin may be implementing a method to access the frame rate more directly.
    # Retrieving it from the displayed status will fail if the frame rate is low enough that no "fps" indicator is displayed in the status
    def getCamFramerate(self):
        # startTime = time.time()
        # startFrame = SharpCap.SelectedCamera.LatestStatus.CapturedFrames 
        # time.sleep(1)   # measure for 1 second
        # endTime = time.time()
        # endFrame = SharpCap.SelectedCamera.LatestStatus.CapturedFrames 
        # fps = (endFrame - startFrame) / (endTime - startTime)
        status=SharpCap.SelectedCamera.LatestStatus.NotificationText
        fps=float(status[status.find(", ")+2:status.find(" fps")]) * 1
        return fps
    
    # framehandler for measuring width - grabs a single frame and looks for first and last transitions, calculates center
    # SharpCap blocks until framehandler returns
    def measureSunFramehandler(self, sender, args):
        frame0 = args.Frame
        # if mean across whole image below threshold, sun is not in frame
        if (frame0.GetStats().Item1 < self.LimbThreshold):
            SharpCap.ShowNotification("*** Sun is not in frame ***", NotificationStatus.Error)
            self.FrameHandlingDone = True
            return

        # scan a 10x100 ROI from left to right to find leading edge of 10 pixel wide window transition
        imgWidth = SharpCap.SelectedCamera.ROI.Width - 10
        startEdge = -1
        x = 0
        while (startEdge<0 and x<imgWidth):
            cutout = frame0.CutROI(Rectangle(x, 0, 10, 100))
            if (cutout.GetStats().Item1 < self.LimbThreshold):
                x += 1
            else:
                startEdge = x

        # scan a 10x100 ROI from right to left to find trailing edge of 10 pixel wide window transition
        endEdge = -1
        x = imgWidth
        while (endEdge<0 and x>=0):
            cutout = frame0.CutROI(Rectangle(x, 0, 10, 100))
            if (cutout.GetStats().Item1 < self.LimbThreshold):
                x -= 1
            else:
                endEdge = x
        
        if (startEdge > 0 and endEdge > 0):
            self.SunWidth = endEdge - startEdge
            self.SunDecenter = (self.SunWidth/2 + startEdge) - (imgWidth+10)/2
        else:
            self.SunWidth = DEFAULT_SUN_WIDTH
            self.sunDecenter = 0
        self.FrameHandlingDone = True

    # Framehandler to detect negative limb transition, check every FrameInterval captured frames
    def acquireFramehandler(self, sender, args):
        if (self.FrameCount == 0):
            try:
                # get stats on center 100x100
                cutout = args.Frame.CutROI(Rectangle(self.ROIX, self.ROIY, 100, 100))
                
                # If still waiting positive transition, check if average is above limb threshold
                if (not self.PositiveSignal):
                    self.PositiveSignal = cutout.GetStats().Item1 > self.LimbThreshold
                # Otherwise check if average is below limb threshold
                elif (not self.EdgePassed):
                    self.EdgePassed = cutout.GetStats().Item1 < self.LimbThreshold
                    
                self.FrameCount = self.FrameInterval      # reset interval counter
            except:
                print("Problem framehandler")
        else:
                self.FrameCount -= 1
        self.FrameHandlingDone = True
        
    # if a bump slew was requested, do 1/4 second slew at the indicated rate. Move the axis not being used for acquisition
    def DoBumpSlew(self):
        SharpCap.Mounts.SelectedMount.MoveAxis(abs(1 - self.AxisToMove), self.BumpSlew)
        time.sleep(0.25)
        SharpCap.Mounts.SelectedMount.MoveAxis(abs(1 - self.AxisToMove), 0)
        self.BumpSlew = 0
        
    # Function to slew in Dec at the given rate until past the limb (brightness drops off below 10%), then for an additional
    # Padded_duration seconds at the Slew_speed_factor rate
    # error out if limb not detected within 30 seconds, and reposition to rough starting position
    # returns True if edge successfully detected, False otherwise
    def SlewPastLimb(self, rate):
        self.AmSlewing = True
        self.EdgePassed = False
        self.PositiveSignal = False
        self.FrameCount = self.FrameInterval
        pad_rate = self.SlewFactor
        math.copysign(pad_rate, rate)   # make padded_slew in same direction

        # set frame handler and start time
        SharpCap.SelectedCamera.FrameCaptured += self.acquireFramehandler
        startTime = time.time()
        
        SharpCap.Mounts.SelectedMount.MoveAxis(self.AxisToMove, rate)
        print(f"Telescope moving at {rate:.2f}x solar speed...")
        while (not self.EdgePassed):     # wait until past limb or 30 seconds passed
            endTime = time.time()
            if (self.TaskAbortFlag):
                return False
            elif (endTime - startTime > 30):
                SharpCap.ShowNotification("\r*** Limb passage not detected within 30 seconds - repositioning mount ***", NotificationStatus.Error)
                SharpCap.Mounts.SelectedMount.MoveAxis(self.AxisToMove, 0)    # Stop slew
                SharpCap.Mounts.SelectedMount.MoveAxis(self.AxisToMove, -rate) # reverse for the same amount of time
                time.sleep(int(endTime - startTime))
                SharpCap.Mounts.SelectedMount.MoveAxis(self.AxisToMove, 0)    # Stop slew, resume tracking
                break
        
        SharpCap.SelectedCamera.FrameCaptured -= self.acquireFramehandler   # unset frame handler
        if (self.EdgePassed):     # if we've successfully detected the negative transition, slew an additional pad and resume tracking
            SharpCap.Mounts.SelectedMount.MoveAxis(self.AxisToMove, pad_rate)
            time.sleep(self.SlewPad)
            SharpCap.Mounts.SelectedMount.MoveAxis(self.AxisToMove, 0)    # Stop slew - rate of 0 restores previous tracking rate
        
        # if a bump slew was requested, do it now
        if (self.BumpSlew != 0):
            self.DoBumpSlew()
        self.AmSlewing = False
        return(self.EdgePassed)
        
    def MeasureSun(self, sender, event):
        # update frame rate
        fps = self.getCamFramerate()
        self.frameRate.Text = f"{fps:.2f}"
        self.doFrameRateChange(None, None)

        # Measure width of bright stripe in image, update sunWidth parameter
        # install framehandler, wait for result
        self.FrameHandlingDone = False
        SharpCap.SelectedCamera.FrameCaptured += self.measureSunFramehandler
        while (not self.FrameHandlingDone):
            pass
        self.sunWidth.Text = str(self.SunWidth)
        self.doSunWidthChange(None, None)
        self.decenter.Text = str(self.SunDecenter)
        
        # uninstall frame handler
        SharpCap.SelectedCamera.FrameCaptured -= self.measureSunFramehandler
        self.CalcScanParams()
        
    def enableGo(self):
       self.goButton.Enabled = True
       self.abortButton.Enabled = False

        
    def enableAbort(self):
       self.abortButton.Enabled = True
       self.goButton.Enabled = False
        
    # do the actual data acquisition
    # There is no good way to interrupt this in SharpCap, should probably be implemented as an asynchronous loop
    # although at least if displayed using ShowDialog, you can Ctrl-C the Python execution
    def asyncDoGo(self, sender, event):
        # enable Abort button, disable Go button
        self.enableAbort()
        self.TaskAbortFlag = False
        Task.Factory.StartNew(self.DoGo)
        
    def DoGo(self):
        # save start coordinates
        self.SavePos()
        
        # slew past edge to starting position
        self.SlewPastLimb(-self.SlewFactor)
        if (self.TaskAbortFlag):
            self.DoAbortTask()
            return
            
        # Main data acquisition loop
        self.progBar.Value = 1
        Num_cycles = int(self.NumCycles)
        for cycle in range(Num_cycles):
            print(f"Cycle {cycle + 1} of {Num_cycles}")
            
            # Start capture
            startTime = time.time()
            SharpCap.SelectedCamera.PrepareToCapture()
            SharpCap.SelectedCamera.RunCapture()
            print("Capture started...")
            self.SlewPastLimb(self.SlewFactor)     # slew until past the limb
            # Stop capture
            SharpCap.SelectedCamera.StopCapture()
            print("Capture stopped.")
            endTime = time.time()
            if (self.TaskAbortFlag):
                self.DoAbortTask()
                return
                
            # If Bidirectional capture on, start the capture
            else:
                if (self.Bidirectional):
                    # capture in reverse direction
                    SharpCap.SelectedCamera.PrepareToCapture()
                    SharpCap.SelectedCamera.RunCapture()
                    print("Reverse capture started...")
                    self.SlewPastLimb(-self.SlewFactor)
                    # Stop capture
                    SharpCap.SelectedCamera.StopCapture()
                    print("Capture stopped.")

                # Otherwise, return at high speed
                else:
                    # return at high speed until past limb, then for pad seconds at forward rate
                    print(f"Returning telescope at {-8*self.SlewFactor:.2f}...")
                    self.SlewPastLimb(-8*self.SlewFactor)
                    print("done")
                    
                # Pause between cycles with a live countdown
                if (self.TaskAbortFlag):
                    self.DoAbortTask()
                    return
                else:
                    print(f"Sleep {self.CycleSleep:.2f} seconds.")
                    time.sleep(float(self.cycleSleep.Text))
            self.progBar.Increment(1)
            
        print("Completed all cycles.")
            
        # Reposition roughly over center of sun
        SharpCap.Mounts.SelectedMount.MoveAxis(self.AxisToMove, self.SlewFactor)
        time.sleep((endTime - startTime)/2)
        SharpCap.Mounts.SelectedMount.MoveAxis(self.AxisToMove, 0)
        self.enableGo()
        
    def RestorePos(self):
        saveRate = SharpCap.Mounts.SelectedMount.SelectedRate
        SharpCap.Mounts.SelectedMount.SelectedRate = Interfaces.AxisRate.ForSiderealRate(32)
        SharpCap.Mounts.SelectedMount.SlewTo(self.SavedCoords)
        SharpCap.Mounts.SelectedMount.SelectedRate = saveRate
    
    def SavePos(self):
        self.SavedCoords = SharpCap.Mounts.SelectedMount.Coordinates
        
    def DoAbort(self, sender, event):
        self.TaskAbortFlag = True

    def DoAbortTask(self):
        # Stop any running capture, stop mount movement, return to saved position
        SharpCap.SelectedCamera.StopCapture()
        SharpCap.Mounts.SelectedMount.MoveAxis(self.AxisToMove, 0)
        self.RestorePos()
        self.AmSlewing = False
        self.BumpSlew = -1
        self.TaskAbortFlag = False
        self.enableGo()

            
    # do bump slews - if mount is currently slewing, set a request, otherwise OK to do the slew ourselves
    def DoBumpL(self, sender, event):
        # slew in RA at the bumpRate for 1 second after acquisition complete
        self.BumpSlew = -self.BumpRate
        if (self.BumpSwap):
            self.BumpSlew *= -1
        if (not self.AmSlewing):
            self.DoBumpSlew()

    def DoBumpLFast(self, sender, event):
        # slew in RA at bumpRate*2 for 1 second after acquisition complete
        self.BumpSlew = -2 * self.BumpRate
        if (self.BumpSwap):
            self.BumpSlew *= -1
        if (not self.AmSlewing):
            self.DoBumpSlew()
        
    def DoBumpR(self, sender, event):
        # slew in negative RA at bumpRate for 1 second after acquisition complete
        self.BumpSlew = self.BumpRate
        if (self.BumpSwap):
            self.BumpSlew *= -1
        if (not self.AmSlewing):
            self.DoBumpSlew()
        
    def DoBumpRFast(self, sender, event):
        # slew in negative RA at bumpRate for 1 second after acquisition complete
        self.BumpSlew = 2 * self.BumpRate
        if (self.BumpSwap):
            self.BumpSlew *= -1
        if (not self.AmSlewing):
            self.DoBumpSlew()
        
    ######################### shutdown #########################
    def BeforeClosing(self, sender, event):
        self.saveSettings()
        
    ######################### build form #########################
    def setupForm(self):
        # informational box
        self.infoLabel = self.addLabel("Start with spectroheliograph positioned over the solar disk\nMeasure solar width at widest point\nSet parameters as desired\nStart acquisition", 12, 2)
        font=self.infoLabel.Font
        self.infoLabel.Font = Font(font.Name, font.Size-2)
        self.FPSInfo = self.addLabel("xxx fps @ blah blah", 22, 56)
        self.FPSInfo.Font = Font(font.Name, font.Size-2)
        
        # input parameters
        self.axisToMove = self.addCheckbox("Slew RA", 30, 82, self.AxisToMove==DEFAULT_AXISTOMOVE, self.doAxisToMoveChange)
        self.addLabel("Number of cycles", 30, 106)
        self.numCycles = self.addTextBox("numCycles", f"{self.NumCycles}", 132, 104 , 45, 20, self.doNumCyclesChange)
        self.bidirectional = self.addCheckbox("Bidirectional", 190, 106, self.Bidirectional, self.doBidirectionalChange)
        self.addLabel("Slew pad (sec)", 30, 130)
        self.slewPad = self.addTextBox("slewPad", f"{self.SlewPad:.1f}", 132, 128, 45, 20, self.doSlewPadChange)
        self.addLabel("Cycle sleep (sec)", 30, 154)
        self.cycleSleep = self.addTextBox("cycleSleep", f"{self.CycleSleep:.1f}", 132, 152, 45, 20, self.doCycleSleepChange)
        self.addLabel("Bump rate", 30, 178)
        self.bumpRate = self.addComboBox("bumpRate", 132, 178, ["1x", "2x", "4x", "8x", "16x"], str(self.BumpRate)+"x", self.doBumpRateChange)
        self.measureSun = self.addButton("Measure Sun", self.MeasureSun, 25, 204)
        self.measureSun.BackColor = Color.Green
        self.sunWidth = self.addTextBox("sunWidth", str(self.SunWidth), 132, 206, 45, 20, self.doSunWidthChange)
        self.addLabel("offset", 190, 210)
        self.decenter = self.addTextBox("decenter", "", 230 , 206, 45, 20, None)
        self.decenter.ReadOnly = True
        self.addLabel("Frame rate", 30, 232)
        self.frameRate = self.addTextBox("Frame rate", "", 132, 230, 45, 20, self.doFrameRateChange)
        
        # control buttons
        self.goButton = self.addButton("Go", self.asyncDoGo, 26, 262)
        self.goButton.BackColor = Color.Green
        self.goButton.Size = Size(124, 28)
        self.abortButton = self.addButton("Abort", self.DoAbort, 196, 262)
        self.abortButton.BackColor = Color.Red
        self.abortButton.Size = Size(124, 28)
        
        # progress bar
        self.progBar = self.addProgressBar("", 29, 294, 288, 10, self.NumCycles)
        
        self.addLabel("swap", 156, 342)
        self.bumpSwap = self.addCheckbox("", 166, 326, self.BumpSwap, self.doBumpSwapChange)
        self.bumpLeft = self.addButton("<-", self.DoBumpL, 68, 324)
        self.bumpLeft.Size = Size(78, 20)
        self.bumpLeftFast = self.addButton("<<", self.DoBumpLFast, 30, 324)
        self.bumpLeftFast.Size = Size(30, 20)
        self.bumpRight = self.addButton("->", self.DoBumpR, 200, 324)
        self.bumpRight.Size = Size(78, 20)
        self.bumpRightFast = self.addButton(">>", self.DoBumpRFast, 288, 324)
        self.bumpRightFast.Size = Size(30, 20)

    def CalcScanParams(self):
        # find top left corner of 100x100 box centered on capture area
        cam=SharpCap.SelectedCamera
        if (cam.ROI.Width<100 or cam.ROI.Height<100):
            SharpCap.ShowNotification("*** Capture ROI must be at least 100x100 pixels ***", NotificationStatus.Error)
            return False
            
        self.ROIX = int(cam.ROI.Width/2 - 50)
        self.ROIY = int(cam.ROI.Height/2 - 50)

        # theoretical required slew rate is calculated assuming need as many lines as width in pixels for 1:1 aspect ratio, and one frame per line
        #   ==> sun_deg / sun_pix = deg/line, multiply by frames (aka lines) per second to obtain required deg/sec
        #   then divide by solar tracking rate of 1/240 deg/sec, which should theoretically result in 120*fps / sunPixWidth
        self.SlewFactor = -(self.FrameRate * 120) / int(self.sunWidth.Text)
        if (self.SlewFactor == 0):
            SharpCap.ShowNotification("*** Frame rate too slow! FPS indicator must be displayed ***", NotificationStatus.Error)
            return False
        cycle_duration = self.SlewPad * 2 + (self.SunWidth/self.FrameRate)
        self.FPSInfo.Text = f"{self.FrameRate:.2f} fps => {abs(self.SlewFactor):.2f}x solar => est cycle duration: {cycle_duration:.2f} sec"
        return True
# end class definition
        
def launch_SHGForm():
    # startup tasks
    global MainForm
    
    # bomb if camera not connected
    if not SharpCap.SelectedCamera:
        SharpCap.ShowNotification("*** Please connect camera before starting SHG scan ***", NotificationStatus.Error)
        return
        
    # connect mount if not already done
    if (not SharpCap.Mounts.SelectedMount.IsConnected):
        SharpCap.Mounts.SelectedMount.Connected = True
            
    # set capture mode to MONO16, SER capture, mount to solar tracking rate
    SharpCap.SelectedCamera.Controls.ColourSpace.Value = "MONO16"
    SharpCap.SelectedCamera.Controls.OutputFormat.Value = "SER file (*.ser)"
    SharpCap.SelectedCamera.LiveView = True
    SharpCap.Mounts.SelectedMount.TrackingRate = Interfaces.TrackingRate.Solar

    MainForm = SHGForm()
    MainForm.StartPosition = FormStartPosition.CenterScreen
    MainForm.TopMost = True
    MainForm.FormClosing += MainForm.BeforeClosing

    fps=MainForm.getCamFramerate()
    MainForm.frameRate.Text = f"{fps:.2f}"
    MainForm.doFrameRateChange(None, None)

    if (MainForm.CalcScanParams()):
        Task.Factory.StartNew(MainForm.ShowDialog)
        time.sleep(0.1)
        MainForm.Activate()
    else:
        MainForm.Close()
    return MainForm

### Main script
# don't create duplicate buttons
if not SharpCap.CustomButtons.Find(lambda x : x.Name == "|   SHG Scan   |"):
    SHGButton = SharpCap.AddCustomButton("|   SHG Scan   |", None, " Perform SHG scanning ", launch_SHGForm)

