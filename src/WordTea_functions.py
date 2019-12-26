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
