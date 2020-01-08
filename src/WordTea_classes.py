#! python3

# Copyright 2017, Dimitrios Stathis, All rights reserved.
# email         : stathis@kth.se, sta.dimitris@gmail.com
# Last edited   : 21/12/2019

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#                                                                         #
#This file is part of WordTea.                                            #
#                                                                         #
#    WordTea is free software: you can redistribute it and/or modify      #
#    it under the terms of the GNU General Public License as published by #
#    the Free Software Foundation, either version 3 of the License, or    #
#    (at your option) any later version.                                  #
#                                                                         #
#    WordTea is distributed in the hope that it will be useful,           #
#    but WITHOUT ANY WARRANTY; without even the implied warranty of       #
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the        #
#    GNU General Public License for more details.                         #
#                                                                         #
#    You should have received a copy of the GNU General Public License    #
#    along with WordTea.  If not, see <https://www.gnu.org/licenses/>.    #
#                                                                         #
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

import sys
import argparse
import os
import re
import comtypes.client
import csv
import docx
import WordTea_functions
from WordTea_functions import *
from docx import Document
from docx.shared import Inches

class reference_list:
    """
    ####################################################################################
    # Class:                                                                           #
    #        reference_list                                                            #
    #                                                                                  #
    # Description:                                                                     #
    #        This class is used to create a set (list) of references. The references   #
    #        can be citations, figures, sections etc. The class includes functions to  #
    #        read and replace references in the text according to specific format.     #
    #        It also outputs warnings for possible formating erros in the final text.  #
    #                                                                                  #
    # TODO Need to write a build list function that builds the list from an input file #
    ####################################################################################
    """

    ################################################################################################
    # Begin : Constructor                                                                          #
    ################################################################################################
    def __init__(self, nm, tg, lb, st=1, pr=None):
        """
        ####################################################################################
        # Function:                                                                        #
        #        reference_list constractor                                                #
        #                                                                                  #
        # Description:                                                                     #
        #        Sets the values for name of the class, tag to be used to search the text  #
        #        and reference style to be replaced in the text                            #
        #                                                                                  #
        # Input Arguments:                                                                 #
        #        nm : Name of the class instance                                           #
        #        tg : Tag to be used for in-text reference                                 #
        #        lb : Label to be searched in-text to create the reference list            #
        #        st : Reference style.Use 1 for normal numbering, 2 for Latin, 3 for small #
        #             letter, 4 for capital letter, default 1.                             #
        #        pr : Parent list, used for linking a hierarchy of sections                #
        #                                                                                  #
        # TODO Check for the correct inputs (string for strings & integer for integer)     #
        ####################################################################################
        """
        self.name = nm
        self.tag = tg
        self.label = lb
        self.style = st
        self.counter = 0
        self.ref_list = list()
        self.parent_count = list()
        if pr is None:
            self.parent = None
        else:
            self.parent = pr
################################################################################################
# End : Constructor                                                                            #
################################################################################################

################################################################################################
# Begin : Build list function                                                                  #
################################################################################################

    def buildList(self, text, v1, v2):
        """
        ####################################################################################
        # Function:                                                                        #
        #        buildList                                                                #
        #                                                                                  #
        # Description:                                                                     #
        #        This function searches the input text for related labels. The labels are  #
        #        related to the class variable label set during the creation of the class. #
        #                                                                                  #
        # Input Arguments:                                                                 #
        #        text    : Text to be searched for related labels                          #
        #        v1      : Verbose level 1                                                 #
        #        v2      : Verbose level 2                                                 #
        #                                                                                  #
        # Return Arguments:                                                                #
        #        text : Function overwrites and returns the text with the label removed    #
        #                                                                                  #
        # TODO : Better verbose                                                            #
        ####################################################################################
        """
        if v2:
            print('###############################################################################')
            print('#            Build list starts Here                                           #')
            print('###############################################################################')
        pr = text
        txt_label = str(self.label)
        if txt_label.lower() in pr.text.lower():
            # Loop added to work with runs (strings with same style)
            tmp = re.split(r'\^', pr.text)
            # This list will contain all the strings that should be deleted from the text
            remove_str = list()
            for pt in tmp:
                tmp_string = txt_label + '{'
                if tmp_string.lower() in pt.lower():
                    remove_str.append(pt)
                    tmp1 = re.split('{|}', pt)
                    tmp1[1] = re.sub(r"\s+", "", tmp1[1])
                    if v2:
                        print(tmp1)
                    # Add the reference to the list and if there is a parrent, register its counter
                    self.ref_list.append(tmp1[1])
                    self.counter += 1
                    if not (self.parent is None):
                        self.parent_count.append(self.parent.counter)
            # Go through the list and delete the labels
            for toDelete in remove_str:
                if v2:
                    print('Remove string')
                    print(toDelete)
                pattern = re.compile(re.escape(str(toDelete)), re.IGNORECASE)
                if v2:
                    print('Before pattern')
                    print(pr.text)
                pr.text = re.sub(pattern,'', pr.text)
                if v2:
                    print('After pattern')
                    print(pr.text)
                pr.text = re.sub(r'\^', '', pr.text)
            del remove_str
        text = pr
        return

################################################################################################
# End : Build list function                                                                    #
################################################################################################

################################################################################################
# Begin : Match and replace function                                                           #
################################################################################################

    def matchNreplace(self, text, v1, v2):
        """
        ####################################################################################
        # Function:                                                                        #
        #        matchNreplace                                                             #
        #                                                                                  #
        # Description:                                                                     #
        #        This function searches the input text for related cross-references.       #
        #        the cross-references are depended to the class variable tag set during    #
        #        the creation of the class.                                                #
        #                                                                                  #
        # Input Arguments:                                                                 #
        #        text    : Text to be searched for related tags                            #
        #        v1      : Verbose level 1                                                 #
        #        v2      : Verbose level 2                                                 #
        #                                                                                  #
        # Return Arguments:                                                                #
        #        text : Function overwrites and returns the text with the label removed    #
        #                                                                                  #
        # TODO : Better verbose                                                            #
        ####################################################################################
        """
        if v2:
            print('###############################################################################')
            print('#            Replace starts Here                                              #')
            print('###############################################################################')

        pr = text
        txt_label = 'ref:'+str(self.tag)
        if txt_label.lower() in pr.text.lower():
            # Loop added to work with runs (strings with same style)
            tmp = re.split(r'\^', pr.text)
            # This list will contain all the strings that should be deleted/replaced from the text
            remove_str = list()
            replace_text = list()
            for pt in tmp:
                tmp_string = txt_label + '{'
                if tmp_string.lower() in pt.lower():
                    remove_str.append(pt)
                    tmp1 = re.split('{|}', pt)
                    tmp1[1] = re.sub(r"\s+", "", tmp1[1]) # tmp1 holds the cross-reference 
                    if v2:
                        print(tmp1)
                    # Go through the ref list and find the cross-reference that matches this one
                    for i in range(len(self.ref_list)):
                        if (tmp1[1].lower() == self.ref_list[i]):
                            replace_text.append(i)
            # Go through the list and delete the labels
            if v2: 
                print('!!!!!!!!!!!!!')
                print('Replace list')
                print(remove_str)
            i = 0
            for toDelete in remove_str:
                if v2:
                    print('##############')
                    print('Replace string')
                    print(toDelete)
                pattern = re.compile(
                    re.escape(str(toDelete)), re.IGNORECASE)
                if v2:
                    print('Before replace pattern')
                    print(pr.text)
                pr.text = re.sub(pattern, str(replace_text[i]+1), pr.text)
                if v2:
                    print('After replace pattern')
                    print(pr.text)
                    print('##############')
                pr.text = re.sub(r'\^', '', pr.text)
                i +=1
            print('!!!!!!!!!!!!!')
            del i
            del replace_text
            del remove_str
        text = pr
        return

################################################################################################
# End : Match and replace function                                                             #
################################################################################################

################################################################################################
# Begin : Match and replace function                                                           #
################################################################################################

    def getParent(self):
        return self.parent

################################################################################################
# End : Match and replace function                                                             #
################################################################################################

################################################################################################
# Begin : Match and replace function                                                           #
################################################################################################

    def getParent_count(self):
        return self.parent_count


################################################################################################
# End : Match and replace function                                                             #
################################################################################################
