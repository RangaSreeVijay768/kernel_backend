from dataclasses import dataclass
from typing import List

# CRC constants
order = 30
polynom = 0x2030B9C7
crcinit = 0x3FFFFFFF
crcxor = 0x3FFFFFFF
direct = 1
refin = 0
refout = 0

# Internal global values
crcmask = (((1 << (order - 1)) - 1) << 1) | 1
crchighbit = 1 << (order - 1)
crcinit_direct = 0
crcinit_nondirect = 0
crctab = [0] * 256

# Bit mask array
ulAdMask = [
    0x0,
    0x1, 0x3, 0x7, 0xF,
    0x1F, 0x3F, 0x7F, 0xFF,
    0x1FF, 0x3FF, 0x7FF, 0xFFF,
    0x1FFF, 0x3FFF, 0x7FFF, 0xFFFF,
    0x1FFFF, 0x3FFFF, 0x7FFFF, 0xFFFFF,
    0x1FFFFF, 0x3FFFFF, 0x7FFFFF, 0xFFFFFF,
    0x1FFFFFF, 0x3FFFFFF, 0x7FFFFFF, 0xFFFFFFF,
    0x1FFFFFFF, 0x3FFFFFFF, 0x7FFFFFFF, 0xFFFFFFFF
]


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
    stDir: List[TagDir]


def reflect(crc: int, bitnum: int) -> int:
    crcout = 0
    for i in range(bitnum):
        if crc & (1 << i):
            crcout |= 1 << (bitnum - 1 - i)
    return crcout


def generate_crc_table():
    for i in range(256):
        crc = i
        if refin:
            crc = reflect(crc, 8)
        crc <<= (order - 8)
        for _ in range(8):
            if crc & crchighbit:
                crc = (crc << 1) ^ polynom
            else:
                crc <<= 1
        if refin:
            crc = reflect(crc, order)
        crc &= crcmask
        crctab[i] = crc


def crcbitbybitfast(p: bytearray, length: int) -> int:
    crc = crcinit_direct
    for byte in p[:length]:
        c = byte
        if refin:
            c = reflect(c, 8)
        for j in range(8)[::-1]:
            bit = crc & crchighbit
            crc <<= 1
            if c & (1 << j):
                bit ^= crchighbit
            if bit:
                crc ^= polynom
    if refout:
        crc = reflect(crc, order)
    crc ^= crcxor
    crc &= crcmask
    return crc


def InsertBits(iStart: int, iNoOfBits: int, pucMsg: bytearray, offset: int, ulDataIn: int) -> int:
    iBitPos = iStart + iNoOfBits

    if iBitPos <= 8:
        iStart = 7 - iStart
        iNBytes = 1
    elif iBitPos <= 16:
        iStart = 15 - iStart
        iNBytes = 2
    else:
        iStart = 31 - iStart
        iNBytes = 4

    # 🔧 Ensure sufficient size in pucMsg
    required_size = offset + iNBytes
    if required_size > len(pucMsg):
        pucMsg.extend([0] * (required_size - len(pucMsg)))

    ulDataBits = 0
    for i in range(iNBytes):
        ulDataBits = (ulDataBits << 8) | pucMsg[offset + i]

    iShiftCount = iStart - iNoOfBits + 1
    ulBitMask = ulAdMask[iNoOfBits]
    ulDataIn &= ulBitMask
    ulDataIn <<= iShiftCount
    ulDataBits &= ~(ulBitMask << iShiftCount)
    ulDataBits |= ulDataIn

    for i in range(iNBytes):
        pucMsg[offset + i] = (ulDataBits >> (8 * (iNBytes - 1 - i))) & 0xFF

    return ulDataBits


def process_taginfo(tag: TagInfo):
    page1 = bytearray(8)
    page2 = bytearray(8)
    total_pages = bytearray(16)

    InsertBits(4, 4, page1, 7, tag.uctypeofTag)
    InsertBits(2, 2, page1, 7, tag.uc_version)
    InsertBits(0, 10, page1, 6, tag.uiUniqueID)
    InsertBits(1, 23, page1, 3, tag.fAbsLoc)
    InsertBits(1, 8, page1, 2, tag.stDir[0].ucTin)
    InsertBits(1, 8, page1, 1, tag.stDir[1].ucTin)

    first_half = tag.stDir[0].StationId & 0x1FF
    second_half = tag.stDir[0].StationId >> 9
    InsertBits(0, 9, page1, 0, first_half)
    InsertBits(1, 7, page2, 7, second_half)

    InsertBits(1, 16, page2, 5, tag.stDir[1].StationId)
    InsertBits(7, 2, page2, 4, tag.stDir[0].secType)
    InsertBits(5, 2, page2, 4, tag.stDir[1].secType)
    InsertBits(2, 3, page2, 4, tag.ucTagPlacement)
    InsertBits(0, 2, page2, 4, tag.AbsoluteLocationReset)
    InsertBits(7, 1, page2, 3, tag.stDir[0].comMark)
    InsertBits(6, 1, page2, 3, tag.stDir[1].comMark)

    for i in range(8):
        total_pages[i] = page1[7 - i]
    for i in range(5):
        total_pages[8 + i] = page2[7 - i]

    crc = crcbitbybitfast(total_pages, 13)
    InsertBits(0, 30, page2, 0, crc)

    return page1, page2, crc


def generate_pages_for_tag(tag: TagInfo):
    generate_crc_table()
    if not direct:
        crc = crcinit
        for _ in range(order):
            bit = crc & crchighbit
            crc <<= 1
            if bit:
                crc ^= polynom
        crc &= crcmask
        globals()['crcinit_direct'] = crc
        globals()['crcinit_nondirect'] = crcinit
    else:
        crc = crcinit
        for _ in range(order):
            bit = crc & 1
            crc >>= 1
            if bit:
                crc ^= polynom
                crc |= crchighbit
        globals()['crcinit_nondirect'] = crc
        globals()['crcinit_direct'] = crcinit

    return process_taginfo(tag)


# ========== MAIN EXECUTION ==========
if __name__ == "__main__":
    if polynom & ~crcmask or crcinit & ~crcmask or crcxor & ~crcmask:
        raise ValueError("CRC parameters are invalid")

    generate_crc_table()

    if not direct:
        crc = crcinit
        for _ in range(order):
            bit = crc & crchighbit
            crc <<= 1
            if bit:
                crc ^= polynom
        crc &= crcmask
        crcinit_direct = crc
        crcinit_nondirect = crcinit
    else:
        crc = crcinit
        for _ in range(order):
            bit = crc & 1
            crc >>= 1
            if bit:
                crc ^= polynom
                crc |= crchighbit
        crcinit_nondirect = crc
        crcinit_direct = crcinit

    # Example tag inputs
    tags = [
        TagInfo(
            uctypeofTag=9,
            uc_version=1,
            uiUniqueID=904,
            fAbsLoc=163820,
            ucTagPlacement=1,
            AbsoluteLocationReset=0,
            stDir=[
                TagDir(ucTin=84, StationId=528, comMark=0, secType=2),
                TagDir(ucTin=84, StationId=527, comMark=0, secType=2)
            ]
        ),
        TagInfo(
            uctypeofTag=9,
            uc_version=1,
            uiUniqueID=904,
            fAbsLoc=163824,
            ucTagPlacement=1,
            AbsoluteLocationReset=0,
            stDir=[
                TagDir(ucTin=84, StationId=528, comMark=0, secType=2),
                TagDir(ucTin=84, StationId=527, comMark=0, secType=2)
            ]
        ),
        TagInfo(
            uctypeofTag=9,
            uc_version=1,
            uiUniqueID=1,
            fAbsLoc=954942,
            ucTagPlacement=0,
            AbsoluteLocationReset=0,
            stDir=[
                TagDir(ucTin=237, StationId=37112, comMark=0, secType=2),
                TagDir(ucTin=237, StationId=37112, comMark=0, secType=2)
            ]
        )
    ]

    for idx, tag in enumerate(tags, 1):
        page1, page2, crc = process_taginfo(tag)
        print(f"\nProcessing Tag {idx} (UniqueID: {tag.uiUniqueID}):")
        print("Page1:", ' '.join(f"{b:02x}" for b in page1))
        print("Page2:", ' '.join(f"{b:02x}" for b in page2))
        print(f"CRC: {crc:08x}")
