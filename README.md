# PPT NDI

## INTRODUCTION
PPT NDI transfers PowerPoint presentations via NDI technology released by NewTek. Thanks to the transparency support, it can be also used as a character generator.

The latest version is PPT NDI (20191117).

## BASIC USAGE
1. Download PPT NDI from https://github.com/ykhwong/ppt-ndi/releases
2. If you have downloaded the installer, execute the pptndi_setup.exe. After the installation, click the Start button and select "PPT NDI".
3. If you have downloaded the 7z file, decompress it and simply run the executable "ppt-ndi.exe"
4. Please select either one of the following modes.

* PowerPoint SlideShow Mode
* Classic Mode

![PPT NDI Mode](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_mode.png)

### PowerPoint SlideShow Mode
Select PowerPoint SlideShow Mode to display the NDI Preview window. Please make sure to disable the "Include background" option to enable transparency.

![PPT NDI Mode: SlideShow](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_slideshow_integration.png)

Open a PowerPoint presentation and start the slide show by pressing Alt-F5 or Alt-Shift-F5.
![PPT NDI SlideShow: SlideShow](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_slideshow_integration2.png)

### Classic Mode
Select Classic Mode to display the dedicated graphical user interface window.

![PPT NDI Screenshot](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_sshot.png)

Click the Open icon to load a Powerpoint file. Please make sure to disable "Include background" option to enable transparency.

## TESTING AND INTEGRATION

The features of PPT NDI can be tested with NDI Studio Monitor. Any NDI-compatible software or hardware can be also used for integration.

![PPT NDI Studio Monitor Integration](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_vmix_example.png)

* obs-ndi users must set the latency mode to Low in the Properties for NDI Source window.
* High-end users can enable high performance mode in the PPT-NDI configuration.


## OTHER NOTES

### Slide resolution
For the best result, please set the slide resolution properly.

1. Open PowerPoint, and go to Design - Slide Size - Custom Slide Size.

![Custom Slide Size 1](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_slide_set_size1.png)

2. Set the width and height manually. For example, if you prefer Full HD, set the width to 1920px and height to 1080px. The px will be automatically converted to either cm or inch depending on the system locale.

![Custom Slide Size 2](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_slide_set_size2.png)

### Configuration

Right-click the system tray icon and click Configure to configure the PPT NDI.

![System tray 1](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_systray1.png)

Here, you can assign system-wide hotkeys. Additionally, you can set PPT  NDI to be minimized in the system tray on startup.

![System tray 2](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_systray2.png)


### Command-Line Options

```sh
  ppt-ndi.exe  [--slideshow] [--classic]
     [--slideshow] : SlidShow Mode
       [--classic] : Classic Mode
```

### Sample

Please also download and open the sample.pptx file to see some examples of the lower thirds:

https://github.com/ykhwong/ppt-ndi/raw/master/resources/sample.pptx

## REQUIREMENT
* Microsoft PowerPoint
* Microsoft Windows 7 or higher (x86-64 only)

## SEE ALSO
* https://www.newtek.com/ndi/
* https://www.newtek.com/ndi/tools/
