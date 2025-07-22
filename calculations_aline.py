import pandas as pd
from typing import Tuple
from dataclasses import dataclass

@dataclass
class TagDir:
    ucTin: int

@dataclass
class TagInfo:
    uctypeofTag: int
    uc_version: int
    uiUniqueID: int
    fAbsLoc: int
    stDir: list[TagDir]
    AdjLine1_tin: int
    AdjLine2_tin: int
    AdjLine3_tin: int
    AdjLine4_tin: int
    AdjLine5_tin: int
    ucTagDuplication: int

@dataclass
class TagEncodedResult:
    page_x: bytes
    page_y: bytes
    crc: int

order = 30
polynom = 0x2030B9C7
crcinit = 0x3FFFFFFF
crcxor = 0x3FFFFFFF
direct = 1
refin = 0
refout = 0

crcmask = ((((1 << (order - 1)) - 1) << 1) | 1)
crchighbit = 1 << (order - 1)
crctab = [0] * 256

def generate_crc_table():
    for i in range(256):
        crc = i
        if refin:
            crc = reflect(crc, 8)
        crc <<= order - 8
        for j in range(8):
            bit = crc & crchighbit
            crc <<= 1
            if bit:
                crc ^= polynom
        if refin:
            crc = reflect(crc, order)
        crc &= crcmask
        crctab[i] = crc

def reflect(crc: int, bitnum: int) -> int:
    crcout = 0
    for i in range(bitnum - 1, -1, -1):
        if crc & (1 << i):
            crcout |= 1 << (bitnum - 1 - i)
    return crcout

# def crcbitbybitfast(data: bytes, length: int) -> int:
#     crc = crcinit
#     for i in range(length):
#         c = data[i]
#         if refin:
#             c = reflect(c, 8)
#         for j in range(0x80, 0, -1):
#             bit = crc & crchighbit
#             crc <<= 1
#             if c & j:
#                 bit ^= crchighbit
#             if bit:
#                 crc ^= polynom
#     if refout:
#         crc = reflect(crc, order)
#     crc ^= crcxor
#     crc &= crcmask
#     return crc



def crc30_cdma(data: bytes, length: int) -> int:
    crc = 0x3FFFFFFF
    for i in range(length):
        crc ^= (data[i] << 22)
        for bit in range(8):
            if crc & 0x20000000:
                crc = (crc << 1) ^ 0x2030B9C7
            else:
                crc <<= 1
            crc &= 0x3FFFFFFF
    crc ^= 0x3FFFFFFF
    return crc & 0x3FFFFFFF


def insert_bits(start: int, nbits: int, buf: bytearray, data: int, offset: int) -> int:
    iBitPos = start + nbits
    if iBitPos <= 8:
        iNBytes = 1
        iStart = 7 - start
    elif iBitPos <= 16:
        iNBytes = 2
        iStart = 15 - start
    elif iBitPos <= 24:
        iNBytes = 3
        iStart = 23 - start
    else:
        iNBytes = 4
        iStart = 31 - start

    if offset + iNBytes > len(buf):
        raise ValueError(f"bytearray index out of range: offset={offset}, required={iNBytes}, buffer len={len(buf)}")

    ulDataBits = 0
    pucMsg = buf[offset:offset + iNBytes]
    for i in range(iNBytes):
        ulDataBits <<= 8
        ulDataBits |= pucMsg[i]
    iShiftCount = iStart - nbits + 1
    ulBitMask = (1 << nbits) - 1
    data &= ulBitMask
    data <<= iShiftCount
    ulDataBits &= ~(ulBitMask << iShiftCount)
    ulDataBits |= data
    for i in range(iNBytes - 1, -1, -1):
        buf[offset + i] = (ulDataBits >> (8 * (iNBytes - 1 - i))) & 0xFF
    return ulDataBits



def calculate_values(tag: TagInfo) -> TagEncodedResult:
    pagex = bytearray(8)
    pagey = bytearray(8)
    total = bytearray(16)

    insert_bits(4, 4, pagex, tag.uctypeofTag & 0xF, offset=7)
    insert_bits(2, 2, pagex, tag.uc_version & 0x3, offset=7)
    insert_bits(0, 10, pagex, tag.uiUniqueID & 0x3FF, offset=6)
    insert_bits(1, 23, pagex, tag.fAbsLoc & 0x7FFFFF, offset=3)
    insert_bits(1, 8, pagex, tag.stDir[0].ucTin & 0xFF, offset=2)
    insert_bits(1, 8, pagex, tag.stDir[1].ucTin & 0xFF, offset=1)
    insert_bits(1, 8, pagex, tag.AdjLine1_tin & 0xFF, offset=0)
    first_half = tag.AdjLine2_tin & 0x01
    second_half = (tag.AdjLine2_tin >> 1) & 0x7F
    insert_bits(0, 1, pagex, first_half, offset=0)
    insert_bits(1, 7, pagey, second_half, offset=7)
    insert_bits(1, 8, pagey, tag.AdjLine3_tin & 0xFF, offset=6)
    insert_bits(1, 8, pagey, tag.AdjLine4_tin & 0xFF, offset=5)
    insert_bits(1, 8, pagey, tag.AdjLine5_tin & 0xFF, offset=4)
    insert_bits(7, 1, pagey, tag.ucTagDuplication & 0x1, offset=3)  # Y31
    insert_bits(6, 2, pagey, 0, offset=3)  # Y33-Y32

    for i in range(8):
        total[i] = pagex[7 - i]
    for i in range(8, 13):
        total[i] = pagey[7 - (i - 8)]
        
    crcc = crc30_cdma(total[:13], 13)
    insert_bits(0, 30, pagey, crcc, offset=0)
    
    print("---- Debug Info ----")
    print("Type:", tag.uctypeofTag, "/ Ver:", tag.uc_version, "/ ID:", tag.uiUniqueID)
    print("AbsLoc:", tag.fAbsLoc)
    print("Dir1:", tag.stDir[0].ucTin, "Dir2:", tag.stDir[1].ucTin)
    print("AdjLines:", tag.AdjLine1_tin, tag.AdjLine2_tin, tag.AdjLine3_tin, tag.AdjLine4_tin, tag.AdjLine5_tin)
    print("Dup:", tag.ucTagDuplication)
    print("PageX:", pagex.hex())
    print("PageY:", pagey.hex())
    print("Total for CRC:", total[:13].hex())
    print("CRC:", hex(crcc))
    print("--------------------")


    return TagEncodedResult(bytes(pagex), bytes(pagey), crcc)


# def calculate_values(tag: TagInfo) -> TagEncodedResult:
#     pagex = bytearray(8)
#     pagey = bytearray(8)
#     total = bytearray(16)

#     # Ensure 8-bit truncation
#     insert_bits(4, 4, pagex, tag.uctypeofTag & 0xF, offset=7)
#     insert_bits(2, 2, pagex, tag.uc_version & 0x3, offset=7)
#     insert_bits(0, 10, pagex, tag.uiUniqueID & 0x3FF, offset=6)
#     insert_bits(1, 23, pagex, tag.fAbsLoc & 0x7FFFFF, offset=3)
#     insert_bits(1, 8, pagex, tag.stDir[0].ucTin & 0xFF, offset=2)
#     insert_bits(1, 8, pagex, tag.stDir[1].ucTin & 0xFF, offset=1)
#     insert_bits(1, 8, pagex, tag.AdjLine1_tin & 0xFF, offset=0)
#     first_half = tag.AdjLine2_tin & 0x01
#     second_half = (tag.AdjLine2_tin >> 1) & 0x7F
#     insert_bits(0, 1, pagex, first_half, offset=0)
#     insert_bits(1, 7, pagey, second_half, offset=7)
#     insert_bits(1, 8, pagey, tag.AdjLine3_tin & 0xFF, offset=6)
#     insert_bits(1, 8, pagey, tag.AdjLine4_tin & 0xFF, offset=5)
#     insert_bits(1, 8, pagey, tag.AdjLine5_tin & 0xFF, offset=4)

#     # Build total with pagex reversed
#     for i in range(8):
#         total[i] = pagex[7 - i]
#     for i in range(8, 13):
#         total[i] = pagey[7 - (i - 8)]

#     crcc = crc30_cdma(total[:13], 13)
#     insert_bits(0, 30, pagey, crcc, offset=0)

#     return TagEncodedResult(bytes(pagex), bytes(pagey), crcc)

generate_crc_table()
