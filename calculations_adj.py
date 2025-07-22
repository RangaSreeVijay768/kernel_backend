# import pandas as pd
# from typing import Tuple
# from dataclasses import dataclass

# @dataclass
# class TagDir:
#     ucTin: int

# @dataclass
# class TagInfo:
#     uctypeofTag: int
#     uc_version: int
#     uiUniqueID: int
#     fAbsLoc1: int
#     fAbsLoc2: int
#     stDir: list[TagDir]
#     dirResetAbsLoc1: int
#     dirResetAbsLoc2: int
#     locCorrectionType: int
#     reserved: int
#     secTypeNominal: int
#     secTypeReverse: int
#     reserved1: int
#     tagType: int
#     comMarkNominal: int
#     comMarkReverse: int

# @dataclass
# class TagEncodedResult:
#     page_x: bytes
#     page_y: bytes
#     crc: int

# order = 30
# polynom = 0x2030B9C7
# crcinit = 0x3FFFFFFF
# crcxor = 0x3FFFFFFF
# direct = 1
# refin = 0
# refout = 0

# crcmask = ((((1 << (order - 1)) - 1) << 1) | 1)
# crchighbit = 1 << (order - 1)
# crctab = [0] * 256

# def generate_crc_table():
#     for i in range(256):
#         crc = i
#         if refin:
#             crc = reflect(crc, 8)
#         crc <<= order - 8
#         for j in range(8):
#             bit = crc & crchighbit
#             crc <<= 1
#             if bit:
#                 crc ^= polynom
#         if refin:
#             crc = reflect(crc, order)
#         crc &= crcmask
#         crctab[i] = crc

# def reflect(crc: int, bitnum: int) -> int:
#     crcout = 0
#     for i in range(bitnum - 1, -1, -1):
#         if crc & (1 << i):
#             crcout |= 1 << (bitnum - 1 - i)
#     return crcout

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



# def insert_bits(start: int, nbits: int, buf: bytearray, data: int, offset: int) -> int:
#     iBitPos = start + nbits
#     if iBitPos <= 8:
#         iNBytes = 1
#         iStart = 7 - start
#     elif iBitPos <= 16:
#         iNBytes = 2
#         iStart = 15 - start
#     elif iBitPos <= 24:
#         iNBytes = 3
#         iStart = 23 - start
#     else:
#         iNBytes = 4
#         iStart = 31 - start

#     if offset + iNBytes > len(buf):
#         raise ValueError(f"bytearray index out of range: offset={offset}, required={iNBytes}, buffer len={len(buf)}")

#     ulDataBits = 0
#     pucMsg = buf[offset:offset + iNBytes]
#     for i in range(iNBytes):
#         ulDataBits <<= 8
#         ulDataBits |= pucMsg[i]
#     iShiftCount = iStart - nbits + 1
#     ulBitMask = (1 << nbits) - 1
#     data &= ulBitMask
#     data <<= iShiftCount
#     ulDataBits &= ~(ulBitMask << iShiftCount)
#     ulDataBits |= data
#     for i in range(iNBytes - 1, -1, -1):
#         buf[offset + i] = (ulDataBits >> (8 * (iNBytes - 1 - i))) & 0xFF
#     return ulDataBits

# def calculate_values(tag: TagInfo) -> TagEncodedResult:
#     page_x = bytearray(8)
#     page_y = bytearray(8)
#     total = bytearray(16)

#     # Ensure 8-bit truncation where necessary
#     insert_bits(4, 4, page_x, tag.uctypeofTag & 0xF, offset=7)  # X3-X0
#     insert_bits(2, 2, page_x, tag.uc_version & 0x3, offset=7)   # X5-X4
#     insert_bits(0, 10, page_x, tag.uiUniqueID & 0x3FF, offset=6) # X15-X6
#     insert_bits(1, 23, page_x, tag.fAbsLoc1 & 0x7FFFFF, offset=3) # X38-X16
#     insert_bits(1, 8, page_x, tag.stDir[0].ucTin & 0xFF, offset=2) # X46-X39
#     insert_bits(1, 8, page_x, tag.stDir[1].ucTin & 0xFF, offset=1) # X54-X47
#     first_half = tag.fAbsLoc2 & 0x1FF  # Lower 9 bits (X63-X55)
#     second_half = (tag.fAbsLoc2 >> 9) & 0x7FFF  # Upper 15 bits (Y14-Y0)
#     insert_bits(0, 9, page_x, first_half, offset=0)  # X63-X55

#     insert_bits(0, 15, page_y, second_half, offset=6)  # Y14-Y0 (starts at bit 0 of page_y[6])
#     insert_bits(1, 3, page_y, tag.dirResetAbsLoc1 & 0x7, offset=6)  # Y16-Y14
#     insert_bits(4, 3, page_y, tag.dirResetAbsLoc2 & 0x7, offset=5)  # Y19-Y17
#     insert_bits(0, 1, page_y, tag.locCorrectionType & 0x1, offset=5)  # Y20
#     insert_bits(1, 2, page_y, tag.reserved & 0x3, offset=4)  # Y22-Y21
#     insert_bits(7, 2, page_y, tag.secTypeNominal & 0x3, offset=3)  # Y24-Y23
#     insert_bits(5, 2, page_y, tag.secTypeReverse & 0x3, offset=3)  # Y26-Y25
#     insert_bits(2, 4, page_y, tag.reserved1 & 0xF, offset=3)  # Y30-Y27
#     insert_bits(1, 1, page_y, tag.tagType & 0x1, offset=3)  # Y31
#     insert_bits(7, 1, page_y, tag.comMarkNominal & 0x1, offset=2)  # Y32
#     insert_bits(6, 1, page_y, tag.comMarkReverse & 0x1, offset=2)  # Y33

#     # Build total with page_x reversed
#     for i in range(8):
#         total[i] = page_x[7 - i]
#     for i in range(8, 13):
#         total[i] = page_y[7 - (i - 8)]

#     crcc = crcbitbybitfast(total[:13], 13)
#     insert_bits(0, 30, page_y, crcc, offset=0)  # Y63-Y34

#     return TagEncodedResult(bytes(page_x), bytes(page_y), crcc)

# generate_crc_table()






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
    fAbsLoc1: int
    fAbsLoc2: int
    stDir: list[TagDir]
    dirResetAbsLoc1: int
    dirResetAbsLoc2: int
    locCorrectionType: int
    reserved: int
    secTypeNominal: int
    secTypeReverse: int
    reserved1: int
    tagType: int
    comMarkNominal: int
    comMarkReverse: int

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

def reflect(crc: int, bitnum: int) -> int:
    crcout = 0
    for i in range(bitnum - 1, -1, -1):
        if crc & (1 << i):
            crcout |= 1 << (bitnum - 1 - i)
    return crcout

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
    page_x = bytearray(8)
    page_y = bytearray(8)
    total = bytearray(16)

    # Ensure 8-bit truncation where necessary
    insert_bits(4, 4, page_x, tag.uctypeofTag & 0xF, offset=7)  # X3-X0
    insert_bits(2, 2, page_x, tag.uc_version & 0x3, offset=7)   # X5-X4
    insert_bits(0, 10, page_x, tag.uiUniqueID & 0x3FF, offset=6) # X15-X6
    insert_bits(1, 23, page_x, tag.fAbsLoc1 & 0x7FFFFF, offset=3) # X38-X16
    insert_bits(1, 8, page_x, tag.stDir[0].ucTin & 0xFF, offset=2) # X46-X39
    insert_bits(1, 8, page_x, tag.stDir[1].ucTin & 0xFF, offset=1) # X54-X47
    first_half = tag.fAbsLoc2 & 0x1FF  # Lower 9 bits (X63-X55)
    second_half = (tag.fAbsLoc2 >> 9) & 0x7FFF  # Upper 15 bits (Y14-Y0)
    insert_bits(0, 9, page_x, first_half, offset=0)  # X63-X55

    insert_bits(0, 15, page_y, second_half, offset=6)  # Y14-Y0
    insert_bits(1, 3, page_y, tag.dirResetAbsLoc1 & 0x7, offset=6)  # Y16-Y14
    insert_bits(4, 3, page_y, tag.dirResetAbsLoc2 & 0x7, offset=5)  # Y19-Y17
    insert_bits(0, 1, page_y, tag.locCorrectionType & 0x1, offset=5)  # Y20
    insert_bits(1, 2, page_y, tag.reserved & 0x3, offset=4)  # Y22-Y21
    insert_bits(7, 2, page_y, tag.secTypeNominal & 0x3, offset=3)  # Y24-Y23
    insert_bits(5, 2, page_y, tag.secTypeReverse & 0x3, offset=3)  # Y26-Y25
    insert_bits(2, 4, page_y, tag.reserved1 & 0xF, offset=3)  # Y30-Y27
    insert_bits(1, 1, page_y, tag.tagType & 0x1, offset=3)  # Y31
    insert_bits(7, 1, page_y, tag.comMarkNominal & 0x1, offset=2)  # Y32
    insert_bits(6, 1, page_y, tag.comMarkReverse & 0x1, offset=2)  # Y33

    # Build total with page_x reversed
    for i in range(8):
        total[i] = page_x[7 - i]
    for i in range(8, 13):
        total[i] = page_y[7 - (i - 8)]

    crcc = crc30_cdma(total[:13], 13)
    insert_bits(0, 30, page_y, crcc, offset=0)  # Y63-Y34

    return TagEncodedResult(bytes(page_x), bytes(page_y), crcc)