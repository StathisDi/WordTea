#! python3

# Copyright 2017, Dimitrios Stathis, All rights reserved.
# email         : stathis@kth.se, sta.dimitris@gmail.com
# Last edited   : 09/01/2020

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

import re
import comtypes.client
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

    def buildList(self, pr, v1, v2):
        """
        ####################################################################################
        # Function:                                                                        #
        #        buildList                                                                 #
        #                                                                                  #
        # Description:                                                                     #
        #        This function searches the input text for related labels. The labels are  #
        #        related to the class variable label set during the creation of the class. #
        #                                                                                  #
        # Input Arguments:                                                                 #
        #        pr      : Text to be searched for related labels                          #
        #        v1      : Verbose level 1                                                 #
        #        v2      : Verbose level 2                                                 #
        #                                                                                  #
        # Return Arguments:                                                                #
        #        pr      : Function overwrites and returns the text with the label removed #
        #                                                                                  #
        ####################################################################################
        """
        if v2:
            print('# Building list paragraph : \n'+pr.text)
        #pr = text
        txt_label = str(self.label)
        error = False
        if txt_label.lower() in pr.text.lower():
            # Loop added to work with runs (strings with same style)
            inline = pr.runs
            for j in range(len(inline)):
                found = False
                [tmp_text, found] = Match_label(inline, j, found, v2, v1, error)
                # If text is found put it in the list
                if found:
                    if v1:
                        print("Info: Match label final text : " + tmp_text)
                    if v2:
                        print("Paragraph after : \n" + pr.text)
                    tmp = FindLabelInText(tmp_text, self.label, v2, v1)
                    # Store the low case label
                    self.ref_list.append(tmp)
                    self.counter += 1
                    if not (self.parent is None):
                        self.parent_count.append(self.parent.counter)
        return 1

################################################################################################
# End : Build list function                                                                    #
################################################################################################

################################################################################################
# Begin : Match and replace function                                                           #
################################################################################################

    def matchNreplace(self, pr, v1, v2):
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
        #        pr      : Text to be searched for related tags                            #
        #        v1      : Verbose level 1                                                 #
        #        v2      : Verbose level 2                                                 #
        #                                                                                  #
        # Return Arguments:                                                                #
        #        pr      : Function overwrites and returns the text with the label removed #
        #                                                                                  #
        # TODO : Better verbose                                                            #
        ####################################################################################
        """
        if v2:
            print('# Searching references in paragraph : \n'+pr.text)
        #pr = text
        txt_label = ':'+str(self.tag)+'{'
        error = False
        if txt_label.lower() in pr.text.lower():
            # Loop added to work with runs (strings with same style)
            inline = pr.runs
            for j in range(len(inline)):
                found = False
                [tmp_text, found] = Match_Tag(inline, j, found, v2, v1, error)
                # If text is found input the wright number
                if found:
                    if v1:
                        print("Info: Match Tag final text : " + tmp_text)
                    if v2:
                        print("Paragraph after removal : \n" + pr.text)
                    for i in range(len(self.ref_list)):
                        if str(self.ref_list[i]).lower() in tmp_text.lower():
                            if (self.style == 1):
                                if (self.parent is None):
                                    txt = str(1 + i)
                            elif (self.style == 2):
                                if (self.parent is None):
                                    txt = int_to_roman(1 + i)
                            elif (self.style == 3):
                                if (self.parent is None):
                                    txt = int_to_small(1 + i)
                            else:
                                if (self.parent is None):
                                    txt = int_to_cap(1 + i)
                            inline[j].text = txt
                    if v2:
                        print("Paragraph after replace : \n" + pr.text)
        return 1

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

################################################################################################
# Begin : Print list                                                                           #
################################################################################################


    def printList(self):
        print("Reference list of " + self.name)
        print(self.ref_list)
        return 1

################################################################################################
# End : Print list                                                                             #
################################################################################################

################################################################################################
# Begin : Print Parent list                                                                    #
################################################################################################

    def printParentList(self):
        rtn = 0
        if not(self.parent is None):
            print("Parent list of " + self.name)
            print(self.parent_count)
            rtn = 1
        else:
            print("No parent for " + self.name + "!")
            rtn = 0
        return rtn

################################################################################################
# End : Print Parent list                                                                      #
################################################################################################

################################################################################################
# Begin : Print                                                                                #
################################################################################################

    def __str__(self):
        rtn = self.name + "\n"
        rtn += "\t Label is: \t\t\t\t"
        rtn += self.label
        rtn += "\n\t Tag is: \t\t\t\t"
        rtn += self.tag
        rtn += "\n\t The chosen style is: \t\t\t"
        rtn += str(self.style)
        rtn += "\n\t The Number of references in the list: \t"
        rtn += str(self.counter)
        if not(self.parent is None):
            rtn += ("\n\t Parent is: \t\t\t\t" + self.parent.name + "\n")
        return rtn

################################################################################################
# End : Print                                                                                  #
################################################################################################

################################################################################################
# Begin : Search and find parrent hiearchy                                                     #
################################################################################################

    def parrentHier(self):
        rtn = ''
        return rtn

################################################################################################
# End : Search and find parrent hiearchy                                                       #
################################################################################################
