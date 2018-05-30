# -*- coding: utf-8 -*-
"""
Created on Fri May 25 10:52:54 2018

@author: johnsonjm
"""

from enum import Enum, unique

@unique
class Answers(Enum):
    Strongly Agree= 1
    Agree = 2
    Disagree = 3
    Strongly Disagree = 4
    Not Applicable = 5