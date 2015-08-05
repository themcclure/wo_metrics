"""Module to read in and process work order (estimate and actuals) information

Broken up into three parts:
1/ Workorder class, and utility functions (query, filter, sort, etc)
2/ Read the WO info data source(s) and return a list of Workorder objects
3/ Report writing module

workorder: where the Wordorder class and utilties live
"""
__author__ = 'mmcclure'

#__all__ = ['Workorder']

from workorder import Workorder
from load_data import load_from_spreadsheet
from generate_reports import build_flow_report

# TODO: rework this if this whole package is going to be run from the command line... probably best done as a non-package script
