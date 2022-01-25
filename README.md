# PPT NDI
[![Github All Releases](https://img.shields.io/github/downloads/ykhwong/ppt-ndi/total.svg?style=flat-square)]()
[![GitHub release](https://img.shields.io/github/release/ykhwong/ppt-ndi.svg?style=flat-square)](https://GitHub.com/ykhwong/ppt-ndi/releases/)

## INTRODUCTION
PPT NDI transfers PowerPoint presentations via NDI technology released by NewTek. Thanks to the transparency support, it can be also used as a character generator.

It also provides SlideShow monitor functionality that supports alpha transparency and can be integrated into any third-party video players that do not support NDI.

The latest version is PPT NDI v1.0.4.

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
* High-end users can enable high performance mode in the PPT-NDI configuration. Enabling the high-performance mode lets the PPT-NDI send frames continuously, which could burden more performance loads. It may help in the case that third-party applications are not able to properly update the NDI image after switching the videos.

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

### Development

For advanced users who want to build the PPT-NDI, please make sure to install:
* Git for Windows
* Visual Studio 2017 15.2 (26418.1 Preview) or higher
* Python 3
* Node.js 10 or higher

Run the below commands:

```sh
git clone https://github.com/ykhwong/ppt-ndi.git
cd ppt-ndi
npm install --save
npm run build
```

#### Experimental macOS Support

For advanced users who want to build the PPT-NDI for macOS, please make sure to install:
* Git
* Command Line Tools for Xcode (to compile PPTNDI.cpp)
* Python 3
* Node.js 10 or higher
* NDI SDK v5

The macOS build only supports classic mode, and uses the experimental internal renderer which is limited to simple text and images, though doesn't have a dependency on Microsoft PowerPoint.

## REQUIREMENT
* Microsoft PowerPoint
* Microsoft Windows 7 or higher (x86-64 only)

When using non-internal renderer on Windows, PPT NDI depends on Visual Basic for Applications (VBA), a component of Microsoft Office. In a default installation of Office, you'll get the VBA installed automatically. People doing a custom installation of Microsoft Office may sometimes change options to exclude VBA. In that way, the PPT NDI has no access to the PowerPoint components.

1. In Control Panel > Programs and Features, locate Microsoft Office. Right-click it, and select Change.
2. On the next panel, select Add or Remove Features.
3. Under Office Shared Features, set Visual Basic for Applications to Run from My Computer. Click Continue.

## SEE ALSO
* https://www.newtek.com/ndi/
* https://www.newtek.com/ndi/tools/
