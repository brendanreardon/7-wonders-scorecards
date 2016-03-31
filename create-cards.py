#!/usr/bin/python

# Import packages
import pandas as pd
#import numpy as np 
import requests
import reportlab

from StringIO import StringIO
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.platypus import Image

# Add fonts
# we know some glyphs are missing, suppress warnings
import reportlab.rl_config
reportlab.rl_config.warnOnMissingFontGlyphs = 0
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(TTFont('caf-bold','cafeteria-black.ttf'))
pdfmetrics.registerFont(TTFont('caf-black','cafeteria-bold.ttf'))
pdfmetrics.registerFont(TTFont('possum-saltare','PossumSaltareNF.ttf'))

# We input our entire data as an xlsx spreadsheet. 
df_all = pd.ExcelFile(r'7 Wonders Score.xlsx')

# We define columns of interest for later on
cols = ['Military', 'Gold', 'Wonder', 'Civic', 'Commercial', 'Scientific', 'Guilds', 'City']

# Define colors which will be used to color the cards
# Color code by frequency strategy. Tableau 10 Medium.
colors = [(237,102,93),(255,158,74),(168,120,110), # Military, Red. Gold, Orange. Wonder, Brown.
           (114,158,206),(205,204,93),(103,191,92), # Civic, Blue. Commerce, Yellow. Scientific, Green.
                       (173,139,201),(162,162,162)] # Guilds, Purple, Cities, Grey.                         
for i in range(len(colors)):
    r, g, b = colors[i]
    colors[i] = r/256.,g/256.,b/256.

# Define the size of our cards
# What is the size of canvas that you want?
width = 4.75
height = 1.25

# We specify the function that draws our card for us.
def wondercards(c):
# Set background color
    c.setFillColor(background_color) #choose fill colour
    c.rect(0, 0, width*inch, height*inch, fill=1)
    
# Place name
    c.setFillColorRGB(256./256,256./256,256./256) #choose your font colour
    c.setStrokeColorRGB(0,0,0)
    c.setFont("possum-saltare", 24) #choose your font type and font size
    c.drawString(85 - 1*inch,135  - 1*inch, name.upper()) # write your text
    #c.setStrokeColorRGB(1,0,0)
    
# Show mode of the rank
    c.setFont("caf-bold", 24) #choose your font type and font size
    c.setFillColorRGB(255./256,221./256,113./256) #choose your font colour
    c.drawString(380 - 1*inch, 135 - 1*inch, str("{0:.0f}".format(mode_rank[0])))

# 7 Wonders reef
    c.drawImage('img/reef.png', 358 - 1*inch, 115 - 1*inch, width = 0.75*inch, height = 0.75*inch, 
        preserveAspectRatio = True, mask=[0, 2, 0, 2, 0, 2])#, 

# Show variance of the rank
    c.setFont("caf-bold", 10) #choose your font type and font size
    c.setFillColorRGB(256./256,256./256,256./256) #choose your font colour
    c.drawString(338 - 1*inch, 150 - 1*inch, ("var: " + str("{0:.2f}".format(var_rank))))  

# Show normalized rank
    c.setFont("caf-bold", 10) #choose your font type and font size
    c.setFillColorRGB(256./256,256./256,256./256) #choose your font colour
    c.drawString(330 - 1*inch, 125 - 1*inch, ("norm: " + str("{0:.2f}".format(norm_rank))))  

# Came in Frist place!
    c.setFillColorRGB(256./256,255./256,255./256) #choose your font colour
    c.setFont("caf-black", 16) #choose your font type and font size
    c.drawString(90 - 1*inch,105 - 1*inch,("Victories: " + str("{0:.0f}".format(wins)))) # write your text

# And games played!
    c.setFillColorRGB(256./256,255./256,255./256) #choose your font colour
    c.setFont("caf-black", 16) #choose your font type and font size
    c.drawString(90 - 1*inch,85 - 1*inch,("Games played: " + str("{0:.0f}".format(games_played)))) # write your text

# Show the high score
    c.setFillColorRGB(256./256,255./256,255./256) #choose your font colour
    c.setFont("caf-black", 16) #choose your font type and font size
    c.drawString(210 - 1*inch,105 - 1*inch,("High Score: " + str("{0:.0f}".format(high_score)))) # write your text
    
# Show the low score
    c.setFillColorRGB(256./256,255./256,255./256) #choose your font colour
    c.setFont("caf-black", 16) #choose your font type and font size
    c.drawString(210 - 1*inch,85 - 1*inch,("Low Score: " + str("{0:.0f}".format(low_score)))) # write your text   

# Show the average rank
    c.setFillColorRGB(256./256,255./256,255./256) #choose your font colour
    c.setFont("caf-black", 16) #choose your font type and font size
    c.setStrokeColorRGB(0.2,0.5,0.3)
    c.drawString(310 - 1*inch,105 - 1*inch,("Avg Rank: " + str("{0:.2f}".format(avg_rank)))) # write your text
    
# Show the average score
    c.setFillColorRGB(256./256,255./256,255./256) #choose your font colour
    c.setFont("caf-black", 16) #choose your font type and font size
    c.drawString(310 - 1*inch,85 - 1*inch,("Avg Score: " + str("{0:.2f}".format(avg_score)))) # write your text

# Now, we actually generate the cards for each sheet. There is a -1 on our range to not include the last
# common scoring sheet. Nate's sample cannot find a mode, so he is excluded too :L for now!
for i in range(0, len(df_all.sheet_names) - 3):
    name = df_all.sheet_names[i]
    
    df = df_all.parse(i)
    df.dropna(how = 'all', inplace = True)

# Calculate quantities of interest:
# Games played, the mode of rank, average rank, average score, low score, 
# high score.
    #wins = df[df['Rank'] == 1].sum()
    #wins = df['Rank'].value_counts()[1]
    wins = (df['Rank']==1).sum()
    games_played = len(df)
    mode_rank = df['Rank'].mode()
    var_rank = df['Rank'].var()
    avg_rank = df['Rank'].mean()
    avg_score = df['Total Score'].mean()
    low_score = df['Total Score'].min()
    high_score = df['Total Score'].max()
    norm_rank = df['Rank'].divide(df['Total Players'])
    norm_rank = norm_rank.mean()

# Calculate which discipline most frequently provides player with the highest
# fraction of overall points. We place an if statement in case the mode is not found.
# This would occur for places when they have not had an avenue grant them the most fraction
# of points more than once.
    freq_strat = df[cols].divide(df['Total Score'], axis = 'index').idxmax(axis = 1).mode()
    if len(freq_strat) == 0:
	   freq_strat = df[cols].divide(df['Total Score'], axis = 'index').idxmax(axis = 1)

# We pick our background color based on this frequent strategy 
    if freq_strat[0] == 'Military':
        background_color = colors[0]
    elif freq_strat[0] == 'Gold':
        background_color = colors[1]
    elif freq_strat[0] == 'Wonder':
        background_color = colors[2]
    elif freq_strat[0] == 'Civic':
        background_color = colors[3]
    elif freq_strat[0] == 'Commercial':
        background_color = colors[4]
    elif freq_strat[0] == 'Scientific':
        background_color = colors[5]
    elif freq_strat[0] == 'Guilds':
        background_color = colors[6]
    elif freq_strat[0] == 'City':
        background_color = colors[7]

    c = canvas.Canvas(r'_Cards/' + name + ".pdf")
    c.setPageSize((width * inch, height * inch))
    print('Printing ' + str(name))

    wondercards(c)
    c.showPage()
    c.save()
# Produce one PDF of all of them stacked ontop of one another.
