# Illustrator Paths to G-Code

This script extracts paths from Adobe Illustrator and creates G-Code for plotting. As of today, this is used to turn an
Ender-3 V2 into a plotter.

## Usage

- Download the script and adjust the configuration values at the top as needed.
- Open the design to be plotted in Illustrator. Make sure that all elements to be plotted are actual paths. If you want
  to plot text, convert it to a path via *Type > Create Outlines*.
- Select and run the script via *File > Scripts > Other Script...*.
- You will be prompted for a location to store the G-Code file to. After selecting a path, the G-Code will be created.

## Development Setup

Adobe provides an [extension to Visual Studio Code](https://marketplace.visualstudio.com/items?itemName=Adobe.extendscript-debug)
that allows to run scripts and attach a debugger. This repository contains the necessary launch configuration.
