#!/usr/bin/python

# Import packages
import pandas as pd
import numpy as np 

from StringIO import StringIO
import requests

import reportlab 
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch

# We input our data as necessary
df_all = pd.ExcelFile(r'7 Wonders Score.xlsx', skipfooter = 18)

for i in range(0, len(df_all.sheet_names) - 2):
    name = df_all.sheet_names[i]
    # name = fix name format 
    
    df = df_all.parse(i)
    df.dropna(how = 'all', inplace = True)
    # We input our datas as necessary
#df = pd.read_csv(r'7 Wonders Score - Iosefa.tsv', engine = 'python', 
#	sep = '\t',	skipfooter = 16)
#df.dropna(how = 'all', inplace = True)

# Calculate quantities of interest:
# Games played, the mode of rank, average rank, average score, low score, 
# high score.
    games_played = len(df)
    mode_rank = df['Rank'].mode()
    avg_rank = df['Rank'].mean()
    avg_score = df['Total Score'].mean()
    low_score = df['Total Score'].min()
    high_score = df['Total Score'].max()

# Calculate which discipline most frequently provides player with the highest
# fraction of overall points. 
    cols = ['Military', 'Gold', 'Wonder', 'Civic', 'Commercial', 'Scientific', 'Guilds', 'City']
    freq_strat = df[cols].divide(df['Total Score'], axis = 'index').idxmax(axis = 1).mode()
# If there is a tie, the mode won't be found.
    if len(freq_strat) == 0:
	   freq_strat = df[cols].divide(df['Total Score'], axis = 'index').idxmax(axis = 1)
#print(freq_strat)

# Color code by frequency strategy. Tableau 10 Medium.
    colors = [(237,102,93),(255,158,74),(168,120,110), # Military, Red. Gold, Orange. Wonder, Brown.
			(114,158,206),(205,204,93),(103,191,92), # Civic, Blue. Commerce, Yellow. Scientific, Green.
						(173,139,201),(162,162,162)] # Guilds, Purple, Cities, Grey. 		                 
    for i in range(len(colors)):
        r, g, b = colors[i]
        colors[i] = r/256.,g/256.,b/256.

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

# What is the size of canvas that you want?
    width = 4.75
    height = 1.25

    def wondercards(c):

    #Third Example
        c.setFillColor(background_color) #choose fill colour
        c.rect(0, 0, width*inch, height*inch, fill=1)
    
    #First Example - Name
        c.setFillColorRGB(256./256,256./256,256./256) #choose your font colour
        c.setFont("Helvetica", 24) #choose your font type and font size
        c.drawString(85 - 1*inch,135  - 1*inch, name) # write your text
    
    #First Example - Mode Rank
        c.setFillColorRGB(255./256,221./256,113./256) #choose your font colour
        c.setFont("Helvetica", 24) #choose your font type and font size
        #c.drawString(300 - 1*inch,135 - 1*inch, str("{0:.0f}".format(mode_rank[0]))) # write your text
        c.drawString(385 - 1*inch, 135 - 1*inch, str("{0:.0f}".format(mode_rank[0])))

    #First Example - Avg Rank
        c.setFillColorRGB(256./256,255./256,255./256) #choose your font colour
        c.setFont("Helvetica", 14) #choose your font type and font size
        c.drawString(300 - 1*inch,105 - 1*inch,(" Avg Rank: " + str("{0:.2f}".format(avg_rank)))) # write your text
    
    #First Example - Avg Score
        c.setFillColorRGB(256./256,255./256,255./256) #choose your font colour
        c.setFont("Helvetica", 14) #choose your font type and font size
        c.drawString(300 - 1*inch,85 - 1*inch,("Avg Score: " + str("{0:.2f}".format(avg_score)))) # write your text
    
    #First Example - High Score
        c.setFillColorRGB(256./256,255./256,255./256) #choose your font colour
        c.setFont("Helvetica", 14) #choose your font type and font size
        c.drawString(200 - 1*inch,105 - 1*inch,("High Score: " + str("{0:.0f}".format(high_score)))) # write your text
    
    #First Example - Low Score
        c.setFillColorRGB(256./256,255./256,255./256) #choose your font colour
        c.setFont("Helvetica", 14) #choose your font type and font size
        c.drawString(200 - 1*inch,85 - 1*inch,(" Low Score: " + str("{0:.0f}".format(low_score)))) # write your text

    # Games played
        c.setFillColorRGB(256./256,255./256,255./256) #choose your font colour
        c.setFont("Helvetica", 14) #choose your font type and font size
        c.drawString(80 - 1*inch,85 - 1*inch,("Games played: " + str("{0:.0f}".format(games_played)))) # write your text
    
    c = canvas.Canvas(r'Cards/' + name + ".pdf")
    c.setPageSize( (4.75*inch, 1.25*inch))

    wondercards(c)
    c.showPage()
    c.save()