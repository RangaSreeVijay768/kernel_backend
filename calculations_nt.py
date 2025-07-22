# # import pandas as pd
# # from typing import Tuple
# # from dataclasses import dataclass

# # @dataclass
# # class TagDir:
# #     ucTin: int
# #     StationId: int
# #     comMark: int
# #     secType: int

# # @dataclass
# # class TagInfo:
# #     uctypeofTag: int
# #     uc_version: int
# #     uiUniqueID: int
# #     fAbsLoc: int
# #     ucTagPlacement: int
# #     AbsoluteLocationReset: int
# #     stDir: list[TagDir]

# # @dataclass
# # class TagEncodedResult:
# #     page_x: bytes
# #     page_y: bytes
# #     crc: int

# # order = 30
# # polynom = 0x2030B9C7
# # crcinit = 0x3FFFFFFF
# # crcxor = 0x3FFFFFFF
# # direct = 1
# # refin = 0
# # refout = 0

# # crcmask = ((((1 << (order - 1)) - 1) << 1) | 1)
# # crchighbit = 1 << (order - 1)
# # crctab = [0] * 256

# # def generate_crc_table():
# #     for i in range(256):
# #         crc = i
# #         if refin:
# #             crc = reflect(crc, 8)
# #         crc <<= order - 8
# #         for j in range(8):
# #             bit = crc & crchighbit
# #             crc <<= 1
# #             if bit:
# #                 crc ^= polynom
# #         if refin:
# #             crc = reflect(crc, order)
# #         crc &= crcmask
# #         crctab[i] = crc

# # def reflect(crc: int, bitnum: int) -> int:
# #     crcout = 0
# #     for i in range(bitnum - 1, -1, -1):
# #         if crc & (1 << i):
# #             crcout |= 1 << (bitnum - 1 - i)
# #     return crcout

# # def crcbitbybitfast(data: bytes, len: int) -> int:
# #     crc = crcinit
# #     for i in range(len):
# #         c = data[i]
# #         if refin:
# #             c = reflect(c, 8)
# #         for j in range(0x80, 0, -1):
# #             bit = crc & crchighbit
# #             crc <<= 1
# #             if c & j:
# #                 bit ^= crchighbit
# #             if bit:
# #                 crc ^= polynom
# #     if refout:
# #         crc = reflect(crc, order)
# #     crc ^= crcxor
# #     crc &= crcmask
# #     return crc




# # def insert_bits(start: int, nbits: int, buf: bytearray, data: int, offset: int) -> int:
# #     iBitPos = start + nbits
# #     if iBitPos <= 8:
# #         iNBytes = 1
# #         iStart = 7 - start
# #     elif iBitPos <= 16:
# #         iNBytes = 2
# #         iStart = 15 - start
# #     elif iBitPos <= 24:
# #         iNBytes = 3
# #         iStart = 23 - start
# #     else:
# #         iNBytes = 4
# #         iStart = 31 - start

# #     if offset + iNBytes > len(buf):
# #         raise ValueError(f"bytearray index out of range: offset={offset}, required={iNBytes}, buffer len={len(buf)}")

# #     ulDataBits = 0
# #     pucMsg = buf[offset:offset + iNBytes]
# #     for i in range(iNBytes):
# #         ulDataBits <<= 8
# #         ulDataBits |= pucMsg[i]

# #     iShiftCount = iStart - nbits + 1
# #     ulBitMask = (1 << nbits) - 1
# #     data &= ulBitMask
# #     data <<= iShiftCount
# #     ulDataBits &= ~(ulBitMask << iShiftCount)
# #     ulDataBits |= data

# #     for i in range(iNBytes - 1, -1, -1):
# #         buf[offset + i] = (ulDataBits >> (8 * (iNBytes - 1 - i))) & 0xFF

# #     return ulDataBits

# # def calculate_values(tag: TagInfo) -> TagEncodedResult:
# #     pagex = bytearray(8)
# #     pagey = bytearray(8)
# #     total = bytearray(16)

# #     # Insert fields into pagex (X0-X63) - unchanged as per request
# #     insert_bits(4, 4, pagex, tag.uctypeofTag, offset=7)  # X3-X0
# #     insert_bits(2, 2, pagex, tag.uc_version, offset=7)   # X5-X4
# #     insert_bits(0, 10, pagex, tag.uiUniqueID, offset=6) # X15-X6
# #     insert_bits(1, 23, pagex, tag.fAbsLoc, offset=3)    # X38-X16
# #     insert_bits(1, 8, pagex, tag.stDir[0].ucTin, offset=2)  # X46-X39
# #     insert_bits(1, 8, pagex, tag.stDir[1].ucTin, offset=1)  # X54-X47
# #     first_half = tag.stDir[0].StationId & 0x1FF  # Lower 9 bits (X63-X55)
# #     second_half = tag.stDir[0].StationId >> 9    # Upper 7 bits (Y6-Y0)
# #     insert_bits(0, 9, pagex, first_half, offset=0)  # X63-X55

# #     # Insert fields into pagey (Y0-Y63)
# #     insert_bits(1, 7, pagey, second_half, offset=7)  # Y6-Y0
# #     insert_bits(1, 16, pagey, tag.stDir[1].StationId, offset=5)  # Y22-Y7
# #     insert_bits(7, 2, pagey, tag.stDir[0].secType, offset=4)     # Y24-Y23
# #     insert_bits(5, 2, pagey, tag.stDir[1].secType, offset=4)     # Y26-Y25
# #     insert_bits(2, 4, pagey, tag.ucTagPlacement, offset=4)       # Y30-Y27
# #     insert_bits(1, 1, pagey, tag.AbsoluteLocationReset, offset=4) # Y31
# #     insert_bits(7, 1, pagey, tag.stDir[0].comMark, offset=3)     # Y32
# #     insert_bits(6, 1, pagey, tag.stDir[1].comMark, offset=3)     # Y33

# #     # Combine pagex and pagey for CRC calculation
# #     for i in range(8):
# #         total[i] = pagex[i]
# #     for i in range(8):
# #         total[8 + i] = pagey[i]

# #     # Calculate CRC (Y63-Y34, 30 bits)
# #     crcc = crcbitbybitfast(total[:13], 13)  # CRC over first 13 bytes
# #     insert_bits(0, 30, pagey, crcc, offset=0)  # Y63-Y34

# #     return TagEncodedResult(bytes(pagex), bytes(pagey), crcc)

# # generate_crc_table()





# import pandas as pd
# from typing import Tuple
# from dataclasses import dataclass

# @dataclass
# class TagDir:
#     ucTin: int
#     StationId: int
#     comMark: int
#     secType: int

# @dataclass
# class TagInfo:
#     uctypeofTag: int
#     uc_version: int
#     uiUniqueID: int
#     fAbsLoc: int
#     ucTagPlacement: int
#     AbsoluteLocationReset: int
#     stDir: list[TagDir]

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

# def crc30_cdma(data: bytes, length: int) -> int:
#     crc = 0x3FFFFFFF
#     for i in range(length):
#         crc ^= (data[i] << 22)
#         for bit in range(8):
#             if crc & 0x20000000:
#                 crc = (crc << 1) ^ 0x2030B9C7
#             else:
#                 crc <<= 1
#             crc &= 0x3FFFFFFF
#     crc ^= 0x3FFFFFFF
#     return crc & 0x3FFFFFFF

# def reflect(crc: int, bitnum: int) -> int:
#     crcout = 0
#     for i in range(bitnum - 1, -1, -1):
#         if crc & (1 << i):
#             crcout |= 1 << (bitnum - 1 - i)
#     return crcout

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
#     pagex = bytearray(8)
#     pagey = bytearray(8)
#     total = bytearray(16)

#     # Insert fields into pagex (X0-X63)
#     insert_bits(4, 4, pagex, tag.uctypeofTag, offset=7)  # X3-X0
#     insert_bits(2, 2, pagex, tag.uc_version, offset=7)   # X5-X4
#     insert_bits(0, 10, pagex, tag.uiUniqueID, offset=6) # X15-X6
#     insert_bits(1, 23, pagex, tag.fAbsLoc, offset=3)    # X38-X16
#     insert_bits(1, 8, pagex, tag.stDir[0].ucTin, offset=2)  # X46-X39
#     insert_bits(1, 8, pagex, tag.stDir[1].ucTin, offset=1)  # X54-X47
#     first_half = tag.stDir[0].StationId & 0x1FF  # Lower 9 bits (X63-X55)
#     second_half = tag.stDir[0].StationId >> 9    # Upper 7 bits (Y6-Y0)
#     insert_bits(0, 9, pagex, first_half, offset=0)  # X63-X55

#     # Insert fields into pagey (Y0-Y63)
#     insert_bits(1, 7, pagey, second_half, offset=7)  # Y6-Y0
#     insert_bits(1, 16, pagey, tag.stDir[1].StationId, offset=5)  # Y22-Y7
#     insert_bits(7, 2, pagey, tag.stDir[0].secType, offset=4)     # Y24-Y23
#     insert_bits(5, 2, pagey, tag.stDir[1].secType, offset=4)     # Y26-Y25
#     insert_bits(2, 4, pagey, tag.ucTagPlacement, offset=4)       # Y30-Y27
#     insert_bits(1, 1, pagey, tag.AbsoluteLocationReset, offset=4) # Y31
#     insert_bits(7, 1, pagey, tag.stDir[0].comMark, offset=3)     # Y32
#     insert_bits(6, 1, pagey, tag.stDir[1].comMark, offset=3)     # Y33

#     # Combine pagex and pagey for CRC calculation
#     for i in range(8):
#         total[i] = pagex[i]
#     for i in range(8):
#         total[8 + i] = pagey[i]

#     # Calculate CRC (Y63-Y34, 30 bits)
#     crcc = crc30_cdma(total[:13], 13)  # CRC over first 13 bytes
#     insert_bits(0, 30, pagey, crcc, offset=0)  # Y63-Y34

#     return TagEncodedResult(bytes(pagex), bytes(pagey), crcc)








import pandas as pd
from typing import Tuple
from dataclasses import dataclass

@dataclass
class TagDir:
    ucTin: int
    StationId: int
    comMark: int
    secType: int

@dataclass
class TagInfo:
    uctypeofTag: int
    uc_version: int
    uiUniqueID: int
    fAbsLoc: int
    ucTagPlacement: int
    AbsoluteLocationReset: int
    stDir: list[TagDir]

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
    pagex = bytearray(8)
    pagey = bytearray(8)

    # Insert bits into pagex
    insert_bits(4, 4, pagex, tag.uctypeofTag, offset=7)   # X3-X0
    insert_bits(2, 2, pagex, tag.uc_version, offset=7)    # X5-X4
    insert_bits(0, 10, pagex, tag.uiUniqueID, offset=6)   # X15-X6
    insert_bits(1, 23, pagex, tag.fAbsLoc, offset=3)      # X38-X16
    insert_bits(1, 8, pagex, tag.stDir[0].ucTin, offset=2)  # X46-X39
    insert_bits(1, 8, pagex, tag.stDir[1].ucTin, offset=1)  # X54-X47

    station_nom = tag.stDir[0].StationId
    station_rev = tag.stDir[1].StationId

    first_half = station_nom & 0x1FF        # X63-X55
    second_half = station_nom >> 9          # Y6-Y0
    insert_bits(0, 9, pagex, first_half, offset=0)

    # Insert partial Y fields (up to Y33) — needed for CRC only
    insert_bits(1, 7, pagey, second_half, offset=7)                 # Y6-Y0
    insert_bits(1, 16, pagey, station_rev, offset=5)               # Y22-Y7
    insert_bits(7, 2, pagey, tag.stDir[0].secType, offset=4)       # Y24-Y23
    insert_bits(5, 2, pagey, tag.stDir[1].secType, offset=4)       # Y26-Y25
    insert_bits(2, 4, pagey, tag.ucTagPlacement, offset=4)         # Y30-Y27
    insert_bits(1, 1, pagey, tag.AbsoluteLocationReset, offset=4)  # Y31
    insert_bits(7, 1, pagey, tag.stDir[0].comMark, offset=3)       # Y32
    insert_bits(6, 1, pagey, tag.stDir[1].comMark, offset=3)       # Y33

    # Build bitstream up to Y33 (94 bits total)
    bitstring = (
        f"{tag.uctypeofTag:04b}"
        f"{tag.uc_version:02b}"
        f"{tag.uiUniqueID:010b}"
        f"{tag.fAbsLoc:023b}"
        f"{tag.stDir[0].ucTin:08b}"
        f"{tag.stDir[1].ucTin:08b}"
        f"{station_nom:016b}"
        f"{station_rev:016b}"
        f"{tag.stDir[0].secType:02b}"
        f"{tag.stDir[1].secType:02b}"
        f"{tag.ucTagPlacement:04b}"
        f"{tag.AbsoluteLocationReset:01b}"
        f"{tag.stDir[0].comMark:01b}"
        f"{tag.stDir[1].comMark:01b}"
    )

    # Pad to full byte (multiple of 8 bits)
    bitstring = bitstring.ljust((len(bitstring) + 7) // 8 * 8, '0')
    data_bytes = int(bitstring, 2).to_bytes(len(bitstring) // 8, byteorder='big')

    # Calculate CRC-30/CDMA
    crcc = crc30_cdma(data_bytes, len(data_bytes))

    # Insert CRC (Y63-Y34)
    insert_bits(0, 30, pagey, crcc, offset=0)

    return TagEncodedResult(bytes(pagex), bytes(pagey), crcc)


