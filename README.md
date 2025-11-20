# Plotter
This repository provides code to plot your data (xy format)

## To install: 
To use the macro you need: -anaconda
						               -matplotlib package 
						               -the macro in itself (the .py file)

To use the macro you need to install anaconda on this website: https://www.anaconda.com/download
anaconda is a distribution of python 

After installation, open an "anaconda prompt" (you can search it on your start tab (onglet démarrer))
in the prompt, write the first time (needed the first time only) : 
	conda install matplotlib
it will install matplotlib, essential to use the macro. Now you can use the macro whevener you want ! 


## General use after installation 
To use the macro, you need to be in the folder where you put the macro's file, you need to give the path to the prompt. For exemple if it is on the file we gave you, you have to write in the prompt: 
	cd C:\Users\Yourname\Nextcloud\Workshop_python

so globally write "cd  Name/of/your/path/where/the/macro/is" (you can copy the path when you are on the right file) 
(we advise you either start you anaconda prompt in the good folder if you use no other code, or just put the file in the base (C:\Users\Yourname) because the conda prompt start there)

When you are in the right path, write in the prompt:
	python Plotter_3.8.py

if you rename the macro, juste replace "Plotter_3.8.py" by the new name you gave. 


We hope it will help you, do not hesitate to use it, or use it as a template for a new macro to create. 
If you see any things not working in the code do not hesitate to tell us!
enjoy plotting! 
Clara and Thomas

PS:For those who want Anaconda Prompt to start directly in their macro folder, simply copy the shortcut.
Then, in the shortcut’s properties under the “Shortcut” tab, set the “Start in” field to the path of the folder where your macro is located.

## Explanation of commands in the gui (graphical user interface)
Note: if one of those parameters is not useful for you, you can just remove it from the gui or put a "#" before

--- Normalize & stacking ---
Normalize = on/off 
normalize by 1 or not your data (not the ref)

offset = 0 /whatever number 
stack your data or if you like put the offset of your choice

refbase = -1 /whatever number
at which y you want your reference to be 

refoffset = 0 /whatever number 
same as "offset" but for the references



--- Title & labels ---
title = write the name of your title here (writing nothing or removing it to have no title)

xlabel = put the label and units of your x axis 

ylabel = put the label and units of your y axis

xlim = 0, 10 
limits of you x-axis to be seen on the plot 

ylim = 0, 50
limits of you y-axis to be seen on the plot 

figsize_cm = 9, 7 
gives the FIGURE (plotting area + framework around) size in cm (not possible to use this in the same time as axes_size_cm)

axes_size_cm = 7,7 
gives the PLOTTING AREA (just the sqaure with your data on, the axis label and the whole framework are not taken into account) size in cm (not possible to use this in the same time as figsize_cm)

magins_cm = 1,1,1,1 #left margin, right, top, bottom
gives the margins outside of your plotting area (so the size of the framework) 



--- Legend & Colors ---
legend = on/off 
if you want or not your legend of data to appear

legendpos = outside
gives the position of your legend (if not present, the computer will choose which one it thinks it is the best)
possible legendpos = upper rith, upper left, uppercenter, center left, center, center right, lower left, lower center, lower right or outside

colormap = rainbow
gives a colormap of your liking to all the data instead of choosing color for each data. all the colormap on matplotlib cheatsheets

legend_labelspacing = 0.1
gives the spacing between the data's label in the legend, if you want them more spaced or more close to each of them

color1 = red 
gives the color red for the first plot of your data. all the color possible are on the matplotlib cheatsheet 

refcolor1 = navy
same as "color1" but for the first plot of reference

line1 = dashed/dotted/dashdot 
gives the style that you want for the 1st plot of your data



--- Font and linewidth Settings ---
font = Arial 
like in Word and Powerpoint, you can choose the font 

textcolor = black
chose the color of text 

linewidth = 1 
gives the width of the line of your data

reflinewidth = 1
same as "linewidth" but for reference plot

label_size = 10
gives the size of writing for the label 

title_size = 12
gives the size of writing for the title

legend_size = 10 
gives the size of xriting for the legend 

square_color = black 
gives the size of the square of your plotting area 

data_bg = white/transparent 
gives a white color for the background of the plotting area or have it transparent 

legendlinewidth = 1 
gives the width for the color part of the legend of your data

legendlinewidthref = 1 
same as before but for the legend of reference


--- Ticks ---
xtick_major = 5/auto/off
gives major ticks on x-axis or remove them. Auto will chose the spacing itself, or you can put a number to choose the spacing that you want 

ytick_major = 2/auto/off
gives major ticks on y-axis or remove them. Auto will chose the spacing itself, or you can put a number to choose the spacing that you want 

xtick_minor = 1.5/auto/off
gives minor ticks (without value) on x-axis or remove them. Auto will chose the spacing itself, or you can put a number to choose the spacing that you want 

ytick_minor = 2/auto/off
gives minor ticks (without value) on y-axis or remove them. Auto will chose the spacing itself, or you can put a number to choose the spacing that you want 


