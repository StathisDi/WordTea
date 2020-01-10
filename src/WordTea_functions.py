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

# TODO : Comment the function
# TODO : Create a util class

import re
import comtypes.client
import docx
from docx import Document
from docx.shared import Inches


# Function to convert integer to latin numera


def int_to_roman(input):
    if not isinstance(input, type(1)):
        raise TypeError(
            "Integer to Latin: Expected integer, got %s" % type(input))
    if not 0 < input < 21:
        raise ValueError("Integer to Latin: Argument must be between 1 and 20")
    ints = (40, 10,  9,   5,  4,   1)
    nums = ('XL', 'X', 'IX', 'V', 'IV', 'I')
    result = []
    for i in range(len(ints)):
        count = int(input / ints[i])
        result.append(nums[i] * count)
        input -= ints[i] * count
    return ''.join(result)

# Function to convert integer to small letters


def int_to_small(input):
    if not isinstance(input, type(1)):
        raise TypeError(
            "Integer to Latin: Expected integer, got %s" % type(input))
    if not 0 < input < 21:
        raise ValueError("Integer to Latin: Argument must be between 1 and 20")
    nums = ('a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j',
            'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't')
    return nums[input]

# Function to convert integer to capital letters


def int_to_cap(input):
    if not isinstance(input, type(1)):
        raise TypeError(
            "Integer to Latin: Expected integer, got %s" % type(input))
    if not 0 < input < 21:
        raise ValueError("Integer to Latin: Argument must be between 1 and 20")
    nums = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
            'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'E', 'S', 'T')
    return nums[input]


def Match_label(inline, j, found, v2, v1, error):
    tmp_txt = ''
    if ('^' in inline[j].text):
        found = True
        tmp_txt = inline[j].text.lower()
        if v2:
            print('First inline: '+tmp_txt)
        k = 1
        flag = True
        while flag:
            pattern = r'\^[^\^]+\^'
            if re.search(pattern, tmp_txt):
                if v2:
                    print('##########################################################')
                    print('             Full Label found:                            ')
                    print(' Label : ' + tmp_txt)
                    print('##########################################################')
                # Remove the label from the inline text
                inline[j].text = re.sub(pattern, '', inline[j].text)
                flag = False  # Exit the loop if you find the full label
            else:
                if '^' in inline[j + k].text:
                    # Remove the initial part of the label from the first inline text
                    toRemove = r'\^[^\^]*$'
                    inline[j].text = re.sub(toRemove, '', inline[j].text)
                    if v2:
                        print('##########################################################')
                        print('                Label is in pieces                        ')
                        print('Inline text: ' + inline[j + k].text + ', k = ' + str(k) + ', j = ' + str(j))
                    split_text = re.split(r'\^', inline[j + k].text)
                    tmp_txt += (split_text[0].lower() + '^')
                    if v2:
                        print('Text list :')
                        print(split_text)
                        print('Final temporary text = ' + tmp_txt)
                    # inline[j + k].text = split_text[1]
                    inline[j + k].text = ''
                    if v2:
                        print('New inline text: ' + inline[j + k].text + ', k = ' + str(k) + ', j = ' + str(j))
                        print('##########################################################')
                    flag = False  # Exit the loop if you complete the label
                else:
                    tmp_txt += inline[j + k].text.lower()
                    temporary_check = j + k
                    if (temporary_check < (len(inline) - 1)):
                        inline[j + k].text = ''
                        k += 1
                    else:
                        print("!!!WARNING!!! Incomplete label :" + inline[j].text)
                        flag = False
                        error = True
                        found = False
    return [tmp_txt, found]


def FindLabelInText(text, label, v2, v1):
    tmp1 = re.split('{|}', text)
    tmp1[1] = re.sub(r"\s+", "", tmp1[1])
    if v1:
        print(tmp1)
        print("Found Label : "+str(tmp1[1]))
    return tmp1[1]
