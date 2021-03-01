# -*- coding: utf-8 -*-
"""
Created on Thu Feb 18 23:22:24 2021

@author: sayers
"""
from src.emailbot import general_move
from time import sleep

while True:
    try:
        general_move()
        sleep(120)
    except:
        print("Encountered an issue. Trying again in 1 minute")
        sleep(60)