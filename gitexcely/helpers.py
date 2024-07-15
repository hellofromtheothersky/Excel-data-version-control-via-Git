import re
import itertools
import time

def find_duplicates(lst :list) -> dict[any, list[int]]:
    duplicate_positions = {}
    
    for i, item in enumerate(lst):
        if item in duplicate_positions:
            duplicate_positions[item].append(i)
        else:
            duplicate_positions[item] = [i]
    
    return {key: value for key, value in duplicate_positions.items() if len(value) > 1}


def column_name_generator():
    # Generate ALPHABET_COL_NAME of lengths 1 to 5 
    ALPHABET_COL_NAME = []
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    # ALPHABET_COL_NAME of length 1
    ALPHABET_COL_NAME.extend(list(alphabet))

    # ALPHABET_COL_NAME of length 2 to 5
    for length in range(2, 3):
        ALPHABET_COL_NAME.extend([''.join(combination) for combination in itertools.product(alphabet, repeat=length)])
    return ALPHABET_COL_NAME


def get_filename_or_lastfoldername(s: str) -> str:
    return re.findall(r'[\\/]{0,1}([^\\\./]*)[\.]{0,1}[\w]*$', s)[0]

def list_str(ls: list[str], pre_sb='- ', post_sb='\n'):
    msg=""
    for l in ls:
        msg+=f"{pre_sb}{l}{post_sb}"
    return msg
