#! python3

# Copyright 2017, Dimitrios Stathis, All rights reserved.
# email         : stathis@kth.se, sta.dimitris@gmail.com
# Version       : 0.1.0
# Last edited   : 13/01/2020

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
import UtilFunctions
from UtilFunctions import *
from docx import Document
from docx.shared import Inches


class referenceList:
    """
    ####################################################################################
    # Class:                                                                           #
    #        referenceList                                                             #
    #                                                                                  #
    # Description:                                                                     #
    #        This class is used to create a set (list) of references. The references   #
    #        can be citations, figures, sections etc. The class includes functions to  #
    #        read and replace references in the text according to specific format.     #
    #        It also outputs warnings for possible formating errors in the final text. #
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
        self.count_list = list()
        self.ref_list = list()
        self.parent_count = list()
        self.oldParent = 0
        self.checkList = list()
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
        # pr = text
        txt_label = '^'+str(self.label)+'{'
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
                    self.checkList.append(0)
                    self.counter += 1
                    if not (self.parent is None):
                        self.parent_count.append(self.parent.counter)
                        if self.counter-1 == 0:
                            self.count_list.append(1)
                            self.parent.printIndexList()
                            self.oldParent = self.parent.getCurrentIndex()
                        else:
                            if self.oldParent == self.parent.getCurrentIndex():
                                temp_count = self.count_list[self.counter - 2] + 1
                                self.count_list.append(temp_count)
                            else:
                                temp_count = 1
                                self.count_list.append(temp_count)
                                self.oldParent = self.parent.getCurrentIndex()
                    else:
                        self.count_list.append(self.counter)
        if(len(self.count_list)!=len(self.ref_list)):
            print("Length of lists does not match!")
            print(self.ref_list)
            print(self.count_list)
            exit()
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
        # pr = text
        txt_label = '`'+str(self.tag)+'{'
        error = False
        if txt_label.lower() in pr.text.lower():
            # Loop added to work with runs (strings with same style)
            inline = pr.runs
            for j in range(len(inline)):
                found = False
                [tmp_text, found] = Match_Tag(inline, j, found, v2, v1, error)
                # If text is found input the right number
                if found:
                    if v1:
                        print("Info: Match Tag final text : " + tmp_text)
                    if v2:
                        print("Paragraph after removal : \n" + pr.text)
                    tagInList = False
                    for i in range(len(self.ref_list)):
                        if str(self.ref_list[i]).lower() in tmp_text.lower():
                            tagInList = True
                            self.checkList[i] = 1
                            if (self.parent is None):
                                txt = formatSelect(self.count_list[i], self.style)
                            else:
                                txt = self.parrentHier(i, v2)
                                txt += formatSelect(self.count_list[i], self.style)
                            inline[j].text += txt
                    if not tagInList:
                        print("############################################!!!WARNING!!!######################################################")
                        print("#Tag or label not found in the list. Either wrong tag or wrong label was used! Check the following paragraph: #\n" + pr.text)
                        print("############################################!!!WARNING!!!######################################################")
                        print("############################################!!!WARNING!!!######################################################")
                        print("#Labels Found  :                                                                                              #")
                        print(tmp_text)
                        print("#Labels in list:                                                                                              #")
                        print(self.ref_list)
                        print("############################################!!!WARNING!!!######################################################")
                    else:
                        if v2:
                            print("Paragraph after replace : \n" + pr.text)
        return 1

################################################################################################
# End : Match and replace function                                                             #
################################################################################################

################################################################################################
# Begin : Get Parent Function                                                                  #
################################################################################################

    def getParent(self):
        return self.parent

################################################################################################
# End : Get parent function                                                                    #
################################################################################################

################################################################################################
# Begin : Get parent count                                                                     #
################################################################################################

    def getParent_count(self):
        return self.parent_count

################################################################################################
# End : Get parent count                                                                       #
################################################################################################

################################################################################################
# Begin : Get current index                                                                    #
################################################################################################

    def getCurrentIndex(self):
        if(self.counter != 0):
            ret = self.count_list[self.counter-1]
        else:
            ret = -1
        return ret

################################################################################################
# End : Get current index                                                                      #
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
# Begin : Print Index list                                                                     #
################################################################################################

    def printIndexList(self):
        print("Index list of " + self.name)
        print(self.count_list)
        return 1

################################################################################################
# End : Print Index list                                                                       #
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
# Begin : Search and find parrent hierarchy                                                    #
################################################################################################

    def parrentHier(self, i, v2):
        rtn = ''
        if not (self.parent is None):
            #print("Parent Count of class " + self.name + " is : " + str(self.parent_count[i]))
            # self.parent.printIndexList()
            #print("Parent index for this count is : "+str(self.parent.count_list[self.parent_count[i]-1]))
            parent_text = self.parent.parrentHier((self.parent_count[i] - 1), v2)
            #print("Parent text in class " + self.name + " is " + parent_text)
            #temp_self_list = self.count_list[i]
            #temporaryText = formatSelect(temp_self_list, self.style)
            #print("Temporary text in class " + self.name + " is " + temporaryText)
            rtn = parent_text + '.'
        else:
            if v2:
                print("No parent in class : " + self.name + ".")
            temp_self_list = self.count_list[i]
            rtn = formatSelect(temp_self_list, self.style)
        if v2:
            print("Text to be returned from the call is : "+rtn)
        return rtn

################################################################################################
# End : Search and find parrent hierarchy                                                      #
################################################################################################

################################################################################################
# Begin : Check reference list                                                                 #
################################################################################################

    def checkRefList(self):
        for i in range(len(self.ref_list)):
            if (self.checkList[i] == 0):
                print("##################!!!Warning!!!##################")
                print("# Label \""+str(self.ref_list[i])+"\" of list "+self.name+"\n"+"# Is not referenced anywhere in the text.")
                print("#################################################")
        return 1

################################################################################################
# End : Check reference list                                                                   #
################################################################################################
