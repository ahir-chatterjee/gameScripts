# -*- coding: utf-8 -*-
"""
Created on Mon Nov 30 16:09:18 2020

@author: Ahir
"""

import xlrd
import xlwt

CREW = "Crew"
IMPOSTER = "Imposter"

readBook = xlrd.open_workbook("UT League Among Us Tracker.xlsx")
gameLog = readBook.sheet_by_name("Live Game Log")

writeBook = xlwt.Workbook()
stats = writeBook.add_sheet('Stat Sheet')

players = {}

for rowNum in range(1,gameLog.nrows):
    #col 0 = date
    #col 1-10 = players
    #col 11 = map
    #col 12-13 = imposters
    #col 14 = crew/imposter win
    
    imposters = [gameLog.cell_value(rowNum,12),gameLog.cell_value(rowNum,13)]
    winners = gameLog.cell_value(rowNum,14)
    
    for colNum in range(1,11): #count statistics for the entire row
        player = gameLog.cell_value(rowNum,colNum)
        
        if(player not in players): #add new player
            players[player] = {"Games Played":1,
                               "Crew Games":0,
                               "Imposter Games":0,
                               "Crew Wins":0,
                               "Imposter Wins":0
                               }
        else: #increment games played
            players[player]["Games Played"] += 1
            
        if(player in imposters): #if player is an imposter, gather their stats
            players[player]["Imposter Games"] += 1
            if(winners == IMPOSTER):
                players[player]["Imposter Wins"] += 1
        else: #if the player is a crewmate, gather their stats
            players[player]["Crew Games"] += 1
            if(winners == CREW):
                players[player]["Crew Wins"] += 1
    

for rowNum, player in enumerate(players, start=1): #output stats into a readable format
    stats.write(rowNum,0,player)
    p = players[player]
    stats.write(rowNum,1,p["Games Played"])
    stats.write(rowNum,2,p["Crew Games"])
    stats.write(rowNum,3,p["Imposter Games"])
    stats.write(rowNum,4,p["Crew Wins"])
    stats.write(rowNum,5,p["Imposter Wins"])
    
writeBook.save("UT League Among Us Stats.xls")