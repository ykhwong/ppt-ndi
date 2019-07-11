# PPT NDI

## INTRODUCTION
PPT NDI transfers PowerPoint presentations via NDI technology released by NewTek. Can be also used as a character generator because it supports transparency.

The latest version is PPT NDI (20190703).

## BASIC USAGE
1. Download PPT NDI from https://github.com/ykhwong/ppt-ndi/releases
2. If you have downloaded the installer, execute the pptndi_setup.exe. After the installation, click Start button and select "PPT NDI".
3. If you have downloaded the 7z file, decompress it and simply run the executable "ppt_ndi.exe"
4. Please select either one of the following modes.

* PowerPoint SlideShow Mode
* Classic Mode

![PPT NDI Mode](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_mode.png)

### PowerPoint SlideShow Mode
In the PowerPoint SlideShow mode, NDI Preview window is provided. Please make sure to disable "Includes background" option to ensure the transparency.

![PPT NDI Mode: SlideShow](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_slideshow_integration.png)

Open PowerPoint presentation and begin a slide show with Alt-F5 or Alt-Shift-F5.
![PPT NDI SlideShow: SlideShow](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_slideshow_integration2.png)

### Classic Mode
In the Classic Mode, a dedicated graphical user interface is displayed.

![PPT NDI Screenshot](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_sshot.png)

Click the Open icon to load a Powerpoint file. Please make sure to disable "Includes background" option to ensure the transparency.

## TESTING AND INTEGRATION

The features of PPT NDI can be tested with NDI Studio Monitor. Any NDI-compatible software or hardware can be also used for the integration.

![PPT NDI Studio Monitor Integration](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_vmix_example.png)

## OTHER NOTES

### Slide resolution
For the best result, please set the slide resolution properly.

1. Open PowerPoint, and go to Design - Slide Size - Custom Slide Size.

![Custom Slide Size 1](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_slide_set_size1.png)

2. Set the width and height manually. For example, if you prefer Full HD, set the width to 1920px and height to 1080px. The px will be automatically converted to either cm or in depending on the system locale.

![Custom Slide Size 2](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_slide_set_size2.png)

### Configuration

Right click the system tray icon and click Configure to configure the PPT NDI.

![System tray 1](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_systray1.png)

Here, you can assign system-wide hot keys. Additionally, you can set up to minimize PPT NDI to system tray when startup.

![System tray 2](https://raw.githubusercontent.com/ykhwong/ppt-ndi/master/resources/ppt_ndi_systray2.png)


### Command Line Options

```sh
  ppt_ndi.exe  [--slideshow] [--classic] [--bg]
     [--slideshow] : SlidShow Mode
       [--classic] : Classic Mode
            [--bg] : Run SlideShow Mode as background (deprecated)
```

### Sample

Please also refer to the sample.pptx that demonstrates the lower thirds:

https://github.com/ykhwong/ppt-ndi/raw/master/resources/sample.pptx

## REQUIREMENT
* Microsoft PowerPoint
* Microsoft Windows 7 or higher (x86-64 only)

## SEE ALSO
* https://www.newtek.com/ndi/
* https://www.newtek.com/ndi/tools/
