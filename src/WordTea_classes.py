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
    def __init__(self, nm, tg, lb, st, pr=None):
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

    def build_list(self, text, verbose):
        """
        ####################################################################################
        # Function:                                                                        #
        #        build_list                                                                #
        #                                                                                  #
        # Description:                                                                     #
        #        This function searches the input text for related labels. The labels are  #
        #        related to the class variable label set during the creation of the class. #
        #                                                                                  #
        # Input Arguments:                                                                 #
        #        text    : Text to be searche for related labels                           #
        #        verbose : Verbose level                                                   #
        #                                                                                  #
        # Return Arguments:                                                                #
        #        text : Function overwrites and returns the text with the label removed    #
        #                                                                                  #
        ####################################################################################
        """
        self.ref_list.append(text)
        print(self.name)
        print(self.ref_list)
        self.counter = self.counter + 1
        if self.parent is not None:
            self.parent_count.append(self.parent.counter)
        return

################################################################################################
# End : Build list function                                                                    #
################################################################################################

################################################################################################
# Begin : Match and replace function                                                           #
################################################################################################

    def matchNreplace(self, text):
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