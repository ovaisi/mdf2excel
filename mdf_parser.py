#!/usr/bin/env python3
"""
SQL Server 2000/2005/2008 MDF file parser.

Parses .MDF database files without requiring a running SQL Server instance.
Extracts table schemas, data rows, stored procedures, indexes, and foreign keys.

Supports:
  - SQL Server 2000 (version 539)
  - SQL Server 2005 (version 611)
  - SQL Server 2008 (version 655)
  - SQL Server 2008 R2 (version 661)
"""

import argparse
import csv
import json
import sqlite3
import os
import struct
import sys
from collections import defaultdict
from ctypes import LittleEndianStructure, c_char, c_int8, c_int16, c_int32, c_uint8, c_uint16
from datetime import datetime, timedelta


PAGE_SIZE = 8192
HEADER_SIZE = 96

# SQL Server base date for datetime columns
SQL_BASE_DATE = datetime(1900, 1, 1)

# Record status byte flags
STATUS_HAS_NULL_BITMAP = 0x10
STATUS_HAS_VARIABLE = 0x20
STATUS_GHOST_FORWARDED = 0x01
STATUS_FORWARDED = 0x04

# Page types
PAGE_DATA = 1
PAGE_INDEX = 2
PAGE_LOB = 3
PAGE_PFS = 8
PAGE_GAM = 9
PAGE_IAM = 10
PAGE_BOOT = 13

# System object IDs for SQL Server 2000
SYSOBJ_SYSOBJECTS = 1
SYSOBJ_SYSINDEXES = 2
SYSOBJ_SYSCOLUMNS = 3
SYSOBJ_SYSTYPES = 4
SYSOBJ_SYSCOMMENTS = 6
SYSOBJ_SYSREFERENCES = 14

# Type IDs (xtype)
XTYPE_IMAGE = 34
XTYPE_TEXT = 35
XTYPE_UNIQUEID = 36
XTYPE_TINYINT = 48
XTYPE_SMALLINT = 52
XTYPE_INT = 56
XTYPE_SMALLDATETIME = 58
XTYPE_REAL = 59
XTYPE_MONEY = 60
XTYPE_DATETIME = 61
XTYPE_FLOAT = 62
XTYPE_NTEXT = 99
XTYPE_BIT = 104
XTYPE_DECIMAL = 106
XTYPE_NUMERIC = 108
XTYPE_SMALLMONEY = 122
XTYPE_BIGINT = 127
XTYPE_VARBINARY = 165
XTYPE_VARCHAR = 167
XTYPE_BINARY = 173
XTYPE_CHAR = 175
XTYPE_TIMESTAMP = 189
XTYPE_NVARCHAR = 231
XTYPE_NCHAR = 239

# Fixed-size types and their sizes
FIXED_TYPE_SIZES = {
    XTYPE_TINYINT: 1,
    XTYPE_SMALLINT: 2,
    XTYPE_INT: 4,
    XTYPE_BIGINT: 8,
    XTYPE_BIT: 1,
    XTYPE_REAL: 4,
    XTYPE_FLOAT: 8,
    XTYPE_MONEY: 8,
    XTYPE_SMALLMONEY: 4,
    XTYPE_SMALLDATETIME: 4,
    XTYPE_DATETIME: 8,
    XTYPE_TIMESTAMP: 8,
    XTYPE_UNIQUEID: 16,
}

# Variable-length type IDs
VARIABLE_TYPES = {XTYPE_IMAGE, XTYPE_TEXT, XTYPE_NTEXT, XTYPE_VARBINARY,
                  XTYPE_VARCHAR, XTYPE_NVARCHAR}

# Types where length is defined by column (char, nchar, binary)
FIXED_WITH_LENGTH_TYPES = {XTYPE_CHAR, XTYPE_NCHAR, XTYPE_BINARY}

# LOB types stored on separate pages
LOB_TYPES = {XTYPE_IMAGE, XTYPE_TEXT, XTYPE_NTEXT}

TYPE_NAMES = {
    XTYPE_IMAGE: 'image', XTYPE_TEXT: 'text', XTYPE_UNIQUEID: 'uniqueidentifier',
    XTYPE_TINYINT: 'tinyint', XTYPE_SMALLINT: 'smallint', XTYPE_INT: 'int',
    XTYPE_SMALLDATETIME: 'smalldatetime', XTYPE_REAL: 'real', XTYPE_MONEY: 'money',
    XTYPE_DATETIME: 'datetime', XTYPE_FLOAT: 'float', XTYPE_NTEXT: 'ntext',
    XTYPE_BIT: 'bit', XTYPE_DECIMAL: 'decimal', XTYPE_NUMERIC: 'numeric',
    XTYPE_SMALLMONEY: 'smallmoney', XTYPE_BIGINT: 'bigint',
    XTYPE_VARBINARY: 'varbinary', XTYPE_VARCHAR: 'varchar',
    XTYPE_BINARY: 'binary', XTYPE_CHAR: 'char', XTYPE_TIMESTAMP: 'timestamp',
    XTYPE_NVARCHAR: 'nvarchar', XTYPE_NCHAR: 'nchar',
}


class PageHeader(LittleEndianStructure):
    """96-byte page header structure."""
    _pack_ = 1
    _fields_ = (
        ('headerVer', c_int8),
        ('type', c_int8),
        ('typeFlag', c_uint8),
        ('level', c_int8),
        ('flag', c_uint16),
        ('indexId', c_int16),
        ('prevPageId', c_int32),
        ('prevFileId', c_int16),
        ('pminlen', c_int16),
        ('nextPageId', c_int32),
        ('nextFileId', c_int16),
        ('slotCnt', c_int16),
        ('objId', c_int32),
        ('freeCnt', c_int16),
        ('freeData', c_int16),
        ('pageId', c_int32),
        ('fileId', c_int16),
        ('reservedCnt', c_int16),
        ('lsn1', c_int32),
        ('lsn2', c_int32),
        ('lsn3', c_int16),
        ('xactReserved', c_int16),
        ('xdesId2', c_int32),
        ('xdesId1', c_int16),
        ('ghostRecCnt', c_int16),
        ('unknown', c_char * 36),
    )

    def __init__(self):
        super().__init__()
        self.unknown = b'\x00'


class MdfFile:
    """Low-level page and record access for MDF files."""

    def __init__(self, path):
        self.path = path
        self.f = open(path, 'rb')
        self.size = os.path.getsize(path)
        self.page_count = self.size // PAGE_SIZE

    def close(self):
        self.f.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()

    def read_page_header(self, page_num):
        """Read the 96-byte page header for page_num."""
        if page_num >= self.page_count:
            return None
        self.f.seek(page_num * PAGE_SIZE)
        hdr = PageHeader()
        self.f.readinto(hdr)
        return hdr

    def read_page_raw(self, page_num):
        """Read full 8192-byte page."""
        self.f.seek(page_num * PAGE_SIZE)
        return self.f.read(PAGE_SIZE)

    def read_slot_array(self, page_num, slot_count):
        """Read the slot array from end of page. Returns list of offsets."""
        base = page_num * PAGE_SIZE
        slots = []
        for i in range(slot_count):
            self.f.seek(base + PAGE_SIZE - 2 * (i + 1))
            offset = struct.unpack('<H', self.f.read(2))[0]
            slots.append(offset)
        return slots

    def read_record_bytes(self, page_num, slot_offset, max_len=8000):
        """Read raw bytes for a record at given offset within page."""
        base = page_num * PAGE_SIZE
        self.f.seek(base + slot_offset)
        return self.f.read(min(max_len, PAGE_SIZE - slot_offset))

    def iter_pages(self, page_type=None, obj_id=None):
        """Iterate over pages, optionally filtering by type and/or objId."""
        for pn in range(self.page_count):
            hdr = self.read_page_header(pn)
            if hdr is None:
                continue
            if page_type is not None and hdr.type != page_type:
                continue
            if obj_id is not None and hdr.objId != obj_id:
                continue
            yield pn, hdr

    def iter_data_pages(self, obj_id):
        """Iterate over data pages (type=1) for a given objId.

        Filters out orphaned (deallocated) pages: when a table has both
        linked pages (prevPageId/nextPageId != 0) and unlinked pages
        (both == 0), only the linked pages contain current data. The
        unlinked pages are leftovers from UPDATE operations where SQL
        Server wrote new rows to a new page chain but didn't zero out
        the old pages.  If ALL pages are unlinked (heap table with no
        clustered index), they are all kept.
        """
        all_pages = list(self.iter_pages(page_type=PAGE_DATA, obj_id=obj_id))
        if len(all_pages) <= 1:
            yield from all_pages
            return

        linked = []
        orphaned = []
        for pn, hdr in all_pages:
            if hdr.prevPageId != 0 or hdr.nextPageId != 0:
                linked.append((pn, hdr))
            else:
                orphaned.append((pn, hdr))

        if linked and orphaned:
            # Mixed: only yield linked pages (orphaned are stale)
            yield from linked
        else:
            # All linked or all orphaned (heap): yield everything
            yield from all_pages

    def get_version(self):
        """Get the database version (dbi_version) from the boot page.

        The boot page is page 9 for all supported versions. The dbi_version
        lives in its data area; rather than rely on a fixed offset that shifts
        between versions, scan the header region for a known version number.
        Falls back to structural detection if none is found.
        """
        if self.page_count > 9:
            page9 = self.read_page_raw(9)
            for offset in range(96, min(400, PAGE_SIZE), 2):
                val = struct.unpack_from('<H', page9, offset)[0]
                if val in (539, 611, 612, 655, 661):
                    return val

        return self._detect_version_from_pages()

    def _detect_version_from_pages(self):
        """Detect version by analyzing page structure patterns."""
        # SQL 2000 has sysobjects at objId=1
        # SQL 2005+ uses different system table object IDs
        has_obj1 = False
        has_large_negids = False
        for pn in range(min(50, self.page_count)):
            hdr = self.read_page_header(pn)
            if hdr is None:
                continue
            if hdr.type == PAGE_DATA:
                if hdr.objId == 1:
                    has_obj1 = True
                if hdr.objId < -100:
                    has_large_negids = True

        if has_obj1 and not has_large_negids:
            return 539  # SQL 2000
        elif has_large_negids:
            return 655  # SQL 2005+ (approximate)
        return 539  # Default assumption

    def get_db_name(self):
        """Try to extract database name from the file header."""
        version = self.get_version()

        # DB name is at a fixed offset on page 9 (boot page) for all versions
        if self.page_count > 9:
            page9 = self.read_page_raw(9)
            # DB name is at offset 148 from page start as nvarchar (128 chars max)
            name_offset = 148
            name_len = 256  # 128 nvarchar chars
            if name_offset + name_len <= len(page9):
                raw = page9[name_offset:name_offset + name_len]
                try:
                    # Find end of UTF-16LE name: look for first null char or
                    # where the data stops being valid UTF-16LE (high byte != 0)
                    null_pos = len(raw)
                    for i in range(0, len(raw) - 1, 2):
                        if raw[i] == 0 and raw[i + 1] == 0:
                            null_pos = i
                            break
                        # If high byte is non-zero and not a valid char, stop
                        if raw[i + 1] != 0 and raw[i] == 0x20 and raw[i + 1] == 0x20:
                            # Space-padded (common in SQL 2000) — stop here
                            null_pos = i
                            break
                    name = raw[:null_pos].decode('utf-16le', errors='replace').strip()
                    if name:
                        return name
                except Exception:
                    pass

        # SQL 2000: parse first record on page 0
        hdr0 = self.read_page_header(0)
        if hdr0.slotCnt > 0:
            slots = self.read_slot_array(0, hdr0.slotCnt)
            if slots:
                rec = self.read_record_bytes(0, slots[0])
                if len(rec) > 4:
                    offset_to_null = struct.unpack_from('<H', rec, 2)[0]
                    if 4 < offset_to_null < len(rec) - 4:
                        pos = offset_to_null
                        if pos + 2 <= len(rec):
                            ncols = struct.unpack_from('<H', rec, pos)[0]
                            pos += 2
                            nbm = (ncols + 7) // 8
                            pos += nbm
                            if pos + 2 <= len(rec):
                                nvar = struct.unpack_from('<H', rec, pos)[0]
                                pos += 2
                                if nvar > 0 and pos + nvar * 2 <= len(rec):
                                    var_offsets = []
                                    for j in range(nvar):
                                        vo = struct.unpack_from('<H', rec, pos)[0]
                                        var_offsets.append(vo & 0x1FFF)
                                        pos += 2
                                    if var_offsets:
                                        name_bytes = rec[pos:var_offsets[0]]
                                        result = _smart_decode(name_bytes)
                                        if result:
                                            return result

        # Fallback: derive from filename
        basename = os.path.basename(self.path)
        if basename.lower().endswith('_data.mdf'):
            return basename[:-9]
        elif basename.lower().endswith('.mdf'):
            return basename[:-4]
        return basename


def decode_value(xtype, raw_bytes, length=None):
    """Decode a raw byte sequence into a Python value based on SQL Server type."""
    if raw_bytes is None or len(raw_bytes) == 0:
        return None

    try:
        if xtype == XTYPE_INT:
            return struct.unpack('<i', raw_bytes[:4])[0]
        elif xtype == XTYPE_SMALLINT:
            return struct.unpack('<h', raw_bytes[:2])[0]
        elif xtype == XTYPE_TINYINT:
            return raw_bytes[0]
        elif xtype == XTYPE_BIGINT:
            return struct.unpack('<q', raw_bytes[:8])[0]
        elif xtype == XTYPE_BIT:
            return bool(raw_bytes[0] & 1)
        elif xtype == XTYPE_FLOAT:
            return struct.unpack('<d', raw_bytes[:8])[0]
        elif xtype == XTYPE_REAL:
            return struct.unpack('<f', raw_bytes[:4])[0]
        elif xtype == XTYPE_MONEY:
            hi, lo = struct.unpack('<iI', raw_bytes[:8])
            val = (hi << 32) | lo
            return val / 10000.0
        elif xtype == XTYPE_SMALLMONEY:
            val = struct.unpack('<i', raw_bytes[:4])[0]
            return val / 10000.0
        elif xtype == XTYPE_DATETIME:
            if len(raw_bytes) < 8:
                return None
            ticks, days = struct.unpack('<Ii', raw_bytes[:8])
            try:
                return SQL_BASE_DATE + timedelta(days=days, milliseconds=ticks * 10 / 3.0)
            except (OverflowError, ValueError):
                return f'<datetime:{days}:{ticks}>'
        elif xtype == XTYPE_SMALLDATETIME:
            if len(raw_bytes) < 4:
                return None
            minutes, days = struct.unpack('<HH', raw_bytes[:4])
            try:
                return SQL_BASE_DATE + timedelta(days=days, minutes=minutes)
            except (OverflowError, ValueError):
                return f'<smalldatetime:{days}:{minutes}>'
        elif xtype in (XTYPE_VARCHAR, XTYPE_CHAR):
            return raw_bytes.decode('ascii', errors='replace').rstrip('\x00').rstrip()
        elif xtype in (XTYPE_NVARCHAR, XTYPE_NCHAR):
            return raw_bytes.decode('utf-16le', errors='replace').rstrip('\x00').rstrip()
        elif xtype in (XTYPE_BINARY, XTYPE_VARBINARY, XTYPE_TIMESTAMP):
            return raw_bytes.hex()
        elif xtype == XTYPE_UNIQUEID:
            if len(raw_bytes) >= 16:
                # SQL Server GUID byte order
                a, b, c = struct.unpack_from('<IHH', raw_bytes, 0)
                d = raw_bytes[8:16]
                return f'{a:08x}-{b:04x}-{c:04x}-{d[:2].hex()}-{d[2:].hex()}'
            return raw_bytes.hex()
        elif xtype in (XTYPE_DECIMAL, XTYPE_NUMERIC):
            if len(raw_bytes) < 1:
                return None
            sign = raw_bytes[0]
            # Remaining bytes are the integer value in little-endian
            val_bytes = raw_bytes[1:]
            val = int.from_bytes(val_bytes, 'little')
            if sign == 0:
                val = -val
            return val  # Note: scale not applied without schema info
        elif xtype in (XTYPE_TEXT, XTYPE_NTEXT, XTYPE_IMAGE):
            # LOB types - data might be inline or on separate pages
            if len(raw_bytes) <= 24:
                # This is likely a LOB pointer, not inline data
                return f'<LOB:{len(raw_bytes)}bytes>'
            # Try to decode inline text
            if xtype == XTYPE_TEXT:
                return raw_bytes.decode('ascii', errors='replace')
            elif xtype == XTYPE_NTEXT:
                return raw_bytes.decode('utf-16le', errors='replace')
            else:
                return f'<image:{len(raw_bytes)}bytes>'
        else:
            return raw_bytes.hex()
    except Exception as e:
        return f'<error:{e}>'


def parse_record(page_data, slot_offset, pminlen=None):
    """Parse a single record from page data at the given offset.

    Returns dict with:
      'status': status byte
      'fixed': bytes of fixed-length data (after 4-byte header)
      'null_bitmap': bytes of null bitmap
      'ncols': number of columns from null bitmap header
      'var_count': number of variable-length columns
      'var_data': list of bytes for each variable-length column
      'raw': full record bytes
    """
    rec = page_data[slot_offset:]
    if len(rec) < 4:
        return None

    status = rec[0]

    # Skip null/empty status records (internal metadata at slot[0])
    if status == 0x00:
        return None

    # Ghost/forwarded records
    if status & STATUS_GHOST_FORWARDED:
        return None

    offset_to_null = struct.unpack_from('<H', rec, 2)[0]
    if offset_to_null < 4 or offset_to_null > PAGE_SIZE:
        return None

    # Note: pminlen is the page's minimum record length, which is a different
    # quantity than offset_to_null (the record's fixed-data end offset). Do not
    # reject records merely because offset_to_null differs from pminlen.

    fixed = rec[4:offset_to_null]

    result = {
        'status': status,
        'fixed': fixed,
        'null_bitmap': b'',
        'ncols': 0,
        'var_count': 0,
        'var_data': [],
        'raw': rec,
    }

    pos = offset_to_null

    # Null bitmap
    if pos + 2 > len(rec):
        return result

    ncols = struct.unpack_from('<H', rec, pos)[0]
    pos += 2
    result['ncols'] = ncols

    nbm_bytes = (ncols + 7) // 8
    if pos + nbm_bytes > len(rec):
        return result

    result['null_bitmap'] = rec[pos:pos + nbm_bytes]
    pos += nbm_bytes

    # Variable-length columns
    if not (status & STATUS_HAS_VARIABLE):
        return result

    if pos + 2 > len(rec):
        return result

    nvar = struct.unpack_from('<H', rec, pos)[0]
    pos += 2
    result['var_count'] = nvar

    if nvar == 0 or nvar > 500:
        return result

    # Read variable column offset array
    var_offsets = []
    for j in range(nvar):
        if pos + 2 > len(rec):
            break
        vo = struct.unpack_from('<H', rec, pos)[0]
        var_offsets.append(vo & 0x1FFF)  # Mask off high bits (flags)
        pos += 2

    # Extract variable column data. Offsets are measured from the record
    # start; the first column's data begins right after the offset array.
    var_data = []
    for i, vo in enumerate(var_offsets):
        if i == 0:
            prev_end = pos  # Data starts after offset array
        else:
            prev_end = var_offsets[i - 1]

        # Offsets are from record start
        col_data = rec[prev_end:vo]
        var_data.append(col_data)

    result['var_data'] = var_data
    return result


def is_null(null_bitmap, col_index):
    """Check if a column is NULL based on the null bitmap. col_index is 0-based."""
    if not null_bitmap or col_index < 0:
        return False
    byte_idx = col_index // 8
    bit_idx = col_index % 8
    if byte_idx >= len(null_bitmap):
        return False
    return bool(null_bitmap[byte_idx] & (1 << bit_idx))


class ColumnDef:
    """Column definition from syscolumns."""
    __slots__ = ('table_id', 'colid', 'name', 'xtype', 'length', 'offset',
                 'is_variable', 'is_nullable', 'prec', 'scale', 'usertype')

    def __init__(self, table_id, colid, name, xtype, length, offset,
                 prec=0, scale=0, usertype=0):
        self.table_id = table_id
        self.colid = colid
        self.name = name
        self.xtype = xtype
        self.length = length
        self.offset = offset
        self.is_variable = xtype in VARIABLE_TYPES
        self.is_nullable = True  # Default; real value comes from status
        self.prec = prec
        self.scale = scale
        self.usertype = usertype

    def type_str(self):
        """Return SQL type string like 'int', 'varchar(50)', 'decimal(10,2)'."""
        base = TYPE_NAMES.get(self.xtype, f'type_{self.xtype}')
        if self.xtype in (XTYPE_CHAR, XTYPE_VARCHAR, XTYPE_BINARY, XTYPE_VARBINARY):
            return f'{base}({self.length})'
        elif self.xtype in (XTYPE_NCHAR, XTYPE_NVARCHAR):
            return f'{base}({self.length // 2})'
        elif self.xtype in (XTYPE_DECIMAL, XTYPE_NUMERIC):
            return f'{base}({self.prec},{self.scale})'
        return base

    def __repr__(self):
        return f'ColumnDef({self.name!r}, {self.type_str()}, colid={self.colid}, offset={self.offset})'


def _smart_decode(data):
    """Detect encoding and decode bytes to string.

    SQL Server stores sysname/nvarchar as UTF-16LE (every other byte is 0x00
    for ASCII text). But some columns (especially syscomments.text in older
    databases) may actually contain plain ASCII despite being typed as nvarchar.
    """
    if not data or len(data) == 0:
        return ''

    # Check if this looks like UTF-16LE: for ASCII-range text,
    # every odd byte should be 0x00
    if len(data) >= 4:
        null_count = sum(1 for i in range(1, min(len(data), 20), 2) if data[i] == 0)
        total_pairs = min(len(data), 20) // 2
        if total_pairs > 0 and null_count >= total_pairs * 0.7:
            # Looks like UTF-16LE
            try:
                return data.decode('utf-16le', errors='replace').rstrip('\x00')
            except Exception:
                pass

    # Try ASCII/Latin-1
    try:
        decoded = data.decode('ascii', errors='replace').rstrip('\x00')
        if decoded:
            return decoded
    except Exception:
        pass

    # Fallback: try UTF-16LE anyway
    try:
        return data.decode('utf-16le', errors='replace').rstrip('\x00')
    except Exception:
        return data.hex()


class SystemCatalog:
    """Reads SQL Server system catalog tables to discover schema."""

    # SQL 2005+ system table object IDs (in page headers)
    SYSOBJ_2005_SYSSCHOBJS = 34
    SYSOBJ_2005_SYSCOLPARS = 41
    SYSOBJ_2005_SYSSCALARTYPES = 50
    SYSOBJ_2005_SYSCOMMENTS = 60  # sys.sysobjvalues for proc text

    def __init__(self, mdf):
        self.mdf = mdf
        self.version = mdf.get_version()
        self._is_2005_plus = None
        self._objects = None  # {obj_id: (name, type)}
        self._columns = None  # {table_id: [ColumnDef, ...]}
        self._types = None    # {xtype: name}
        self._comments = None  # {obj_id: [(colid, text), ...]}

    @property
    def is_2005_plus(self):
        """Detect whether this database uses SQL 2005+ catalog format."""
        if self._is_2005_plus is not None:
            return self._is_2005_plus

        # Check if sysobjects (objId=1) has data pages
        has_sysobjects = False
        has_sysschobjs = False
        for pn, hdr in self.mdf.iter_data_pages(SYSOBJ_SYSOBJECTS):
            if hdr.slotCnt > 0:
                has_sysobjects = True
                break
        for pn, hdr in self.mdf.iter_data_pages(self.SYSOBJ_2005_SYSSCHOBJS):
            if hdr.slotCnt > 0:
                has_sysschobjs = True
                break

        # SQL 2000 uses sysobjects at objId=1
        # SQL 2005+ uses sysschobjs at objId=34
        # If we find sysschobjs but not sysobjects with valid type codes, it's 2005+
        if has_sysobjects and not has_sysschobjs:
            self._is_2005_plus = False
        elif has_sysschobjs:
            # Verify sysschobjs has valid object type codes at offset 13
            for pn, hdr in self.mdf.iter_data_pages(self.SYSOBJ_2005_SYSSCHOBJS):
                page_data = self.mdf.read_page_raw(pn)
                slots = self.mdf.read_slot_array(pn, hdr.slotCnt)
                for slot_off in slots:
                    if slot_off < HEADER_SIZE:
                        continue
                    rec = parse_record(page_data, slot_off)
                    if rec and len(rec['fixed']) >= 15:
                        type_bytes = rec['fixed'][13:15]
                        try:
                            t = type_bytes.decode('ascii', errors='replace').strip()
                            if t in ('U', 'P', 'S', 'V', 'PK', 'F', 'D', 'IT', 'SQ'):
                                self._is_2005_plus = True
                                return self._is_2005_plus
                        except Exception:
                            pass
                break
            self._is_2005_plus = has_sysschobjs
        else:
            self._is_2005_plus = False

        return self._is_2005_plus

    def _read_all_records(self, obj_id):
        """Read all records from data pages belonging to obj_id."""
        records = []
        for pn, hdr in self.mdf.iter_data_pages(obj_id):
            if hdr.slotCnt <= 0:
                continue
            page_data = self.mdf.read_page_raw(pn)
            slots = self.mdf.read_slot_array(pn, hdr.slotCnt)
            for slot_off in slots:
                if slot_off < HEADER_SIZE or slot_off >= PAGE_SIZE:
                    continue
                rec = parse_record(page_data, slot_off, hdr.pminlen)
                if rec is not None:
                    records.append(rec)
        return records

    def get_objects(self):
        """Parse sysobjects/sysschobjs. Returns {obj_id: (name, type_char)}."""
        if self._objects is not None:
            return self._objects

        self._objects = {}

        if self.is_2005_plus:
            self._parse_sysschobjs()
        else:
            self._parse_sysobjects()

        return self._objects

    def _parse_sysobjects(self):
        """Parse SQL 2000 sysobjects (objId=1)."""
        records = self._read_all_records(SYSOBJ_SYSOBJECTS)

        for rec in records:
            fixed = rec['fixed']
            if len(fixed) < 6:
                continue

            obj_id = struct.unpack_from('<i', fixed, 0)[0]
            xtype_bytes = fixed[4:6]
            try:
                obj_type = xtype_bytes.decode('ascii', errors='replace').strip()
            except Exception:
                obj_type = '??'

            name = ''
            if rec['var_data'] and rec['var_data'][0]:
                try:
                    name = rec['var_data'][0].decode('utf-16le', errors='replace').rstrip('\x00')
                except Exception:
                    name = _smart_decode(rec['var_data'][0])

            self._objects[obj_id] = (name, obj_type)

    def _parse_sysschobjs(self):
        """Parse SQL 2005+ sys.sysschobjs (objId=34)."""
        records = self._read_all_records(self.SYSOBJ_2005_SYSSCHOBJS)

        for rec in records:
            fixed = rec['fixed']
            if len(fixed) < 15:
                continue

            # sysschobjs fixed layout (pminlen=26):
            # [0:4]   = id (int) — object id
            # [4:8]   = nsid (int)
            # [8]     = nsclass (tinyint)
            # [9:13]  = status (int)
            # [13:15] = type (char(2)) — object type code
            obj_id = struct.unpack_from('<i', fixed, 0)[0]
            type_bytes = fixed[13:15]
            try:
                obj_type = type_bytes.decode('ascii', errors='replace').strip()
            except Exception:
                obj_type = '??'

            name = ''
            if rec['var_data'] and rec['var_data'][0]:
                try:
                    name = rec['var_data'][0].decode('utf-16le', errors='replace').rstrip('\x00')
                except Exception:
                    name = _smart_decode(rec['var_data'][0])

            self._objects[obj_id] = (name, obj_type)

    def get_columns(self):
        """Parse syscolumns/syscolpars. Returns {table_id: [ColumnDef, ...]}."""
        if self._columns is not None:
            return self._columns

        self._columns = defaultdict(list)

        if self.is_2005_plus:
            self._parse_syscolpars()
        else:
            self._parse_syscolumns()

        # Deduplicate columns by (table_id, colid) — MDF slot arrays can
        # contain duplicate pointers to the same record after page splits
        for tid in self._columns:
            seen = {}
            deduped = []
            for col in self._columns[tid]:
                if col.colid not in seen:
                    seen[col.colid] = col
                    deduped.append(col)
            self._columns[tid] = deduped

        # Sort columns by colid within each table
        for tid in self._columns:
            self._columns[tid].sort(key=lambda c: c.colid)

        # For SQL 2005+, compute offsets since they're not stored in syscolpars
        if self.is_2005_plus:
            self._compute_column_offsets()

        return self._columns

    def _parse_syscolumns(self):
        """Parse SQL 2000 syscolumns (objId=3)."""
        records = self._read_all_records(SYSOBJ_SYSCOLUMNS)

        for rec in records:
            fixed = rec['fixed']
            if len(fixed) < 16:
                continue

            table_id = struct.unpack_from('<i', fixed, 0)[0]
            xtype = fixed[4]
            col_length = struct.unpack_from('<h', fixed, 8)[0]
            prec = fixed[10] if len(fixed) > 10 else 0
            scale = fixed[11] if len(fixed) > 11 else 0
            colid = struct.unpack_from('<h', fixed, 12)[0]
            col_offset = struct.unpack_from('<h', fixed, 14)[0]

            name = ''
            if rec['var_data'] and rec['var_data'][0]:
                try:
                    name = rec['var_data'][0].decode('utf-16le', errors='replace').rstrip('\x00')
                except Exception:
                    name = _smart_decode(rec['var_data'][0])
                if not name:
                    name = f'col_{colid}'

            col = ColumnDef(
                table_id=table_id, colid=colid, name=name,
                xtype=xtype, length=col_length, offset=col_offset,
                prec=prec, scale=scale,
            )
            self._columns[table_id].append(col)

    def _parse_syscolpars(self):
        """Parse SQL 2005+ sys.syscolpars (objId=41)."""
        records = self._read_all_records(self.SYSOBJ_2005_SYSCOLPARS)

        for rec in records:
            fixed = rec['fixed']
            if len(fixed) < 19:
                continue

            # syscolpars fixed layout (pminlen=88):
            # [0:4]   = id (int) — table object id
            # [4:6]   = number (smallint)
            # [6:8]   = colid (smallint)
            # [8:10]  = (padding)
            # [10]    = xtype (tinyint)
            # [11]    = utype (tinyint)
            # [12:14] = (padding)
            # [14]    = (flags)
            # [15:17] = length (smallint)
            # [17]    = prec (tinyint)
            # [18]    = scale (tinyint)
            table_id = struct.unpack_from('<i', fixed, 0)[0]
            colid = struct.unpack_from('<h', fixed, 6)[0]
            xtype = fixed[10]
            col_length = struct.unpack_from('<h', fixed, 15)[0]
            prec = fixed[17] if len(fixed) > 17 else 0
            scale = fixed[18] if len(fixed) > 18 else 0

            # Offset will be computed later
            col_offset = 0

            name = ''
            if rec['var_data'] and rec['var_data'][0]:
                try:
                    name = rec['var_data'][0].decode('utf-16le', errors='replace').rstrip('\x00')
                except Exception:
                    name = _smart_decode(rec['var_data'][0])
                if not name:
                    name = f'col_{colid}'

            col = ColumnDef(
                table_id=table_id, colid=colid, name=name,
                xtype=xtype, length=col_length, offset=col_offset,
                prec=prec, scale=scale,
            )
            self._columns[table_id].append(col)

    def _compute_column_offsets(self):
        """Compute record byte offsets for SQL 2005+ columns.

        SQL 2005+ syscolpars doesn't store physical offsets. We compute them
        by ordering fixed-length columns by colid and accumulating sizes.
        Variable-length columns get negative offsets (ordered by colid).
        """
        for tid, cols in self._columns.items():
            offset = 4  # After 4-byte record header
            var_index = 0

            for col in cols:
                if col.xtype in VARIABLE_TYPES or col.xtype in LOB_TYPES:
                    var_index -= 1
                    col.offset = var_index  # Negative for variable columns
                else:
                    col.offset = offset
                    # Determine size for this column
                    if col.xtype in FIXED_TYPE_SIZES:
                        size = FIXED_TYPE_SIZES[col.xtype]
                    elif col.xtype in FIXED_WITH_LENGTH_TYPES:
                        size = col.length
                    elif col.xtype in (XTYPE_DECIMAL, XTYPE_NUMERIC):
                        size = col.length
                    else:
                        size = col.length
                    offset += max(size, 0)

    def get_types(self):
        """Parse systypes/sysscalartypes. Returns {xtype: name}."""
        if self._types is not None:
            return self._types

        self._types = {}

        if self.is_2005_plus:
            records = self._read_all_records(self.SYSOBJ_2005_SYSSCALARTYPES)
            for rec in records:
                fixed = rec['fixed']
                if len(fixed) < 4:
                    continue
                # sysscalartypes: fixed[0] = xtype (matches system_type_id)
                xtype = fixed[0]
                name = ''
                if rec['var_data'] and rec['var_data'][0]:
                    try:
                        name = rec['var_data'][0].decode('utf-16le', errors='replace').rstrip('\x00')
                    except Exception:
                        name = _smart_decode(rec['var_data'][0])
                if name:
                    self._types[xtype] = name
        else:
            records = self._read_all_records(SYSOBJ_SYSTYPES)
            for rec in records:
                fixed = rec['fixed']
                if len(fixed) < 4:
                    continue
                xtype = fixed[0]
                name = ''
                if rec['var_data'] and rec['var_data'][0]:
                    try:
                        name = rec['var_data'][0].decode('utf-16le', errors='replace').rstrip('\x00')
                    except Exception:
                        name = _smart_decode(rec['var_data'][0])
                if name:
                    self._types[xtype] = name

        return self._types

    def get_comments(self):
        """Parse syscomments/sysobjvalues. Returns {obj_id: [(number, colid, text), ...]}."""
        if self._comments is not None:
            return self._comments

        self._comments = defaultdict(list)

        if self.is_2005_plus:
            # SQL 2005+: sysobjvalues (objId=60)
            # fixed layout: valclass(1) + objid(4) + subobjid(4) + ...
            # var[1] = SQL text (ASCII)
            records = self._read_all_records(self.SYSOBJ_2005_SYSCOMMENTS)
            for rec in records:
                fixed = rec['fixed']
                if len(fixed) < 5:
                    continue
                valclass = fixed[0]
                if valclass != 1:  # Only valclass=1 contains proc/view definitions
                    continue
                obj_id = struct.unpack_from('<i', fixed, 1)[0]
                subobjid = struct.unpack_from('<i', fixed, 5)[0] if len(fixed) >= 9 else 0

                text = ''
                if len(rec['var_data']) >= 2 and rec['var_data'][1]:
                    text = _smart_decode(rec['var_data'][1])

                if text:
                    self._comments[obj_id].append((0, subobjid, text))
        else:
            # SQL 2000: syscomments (objId=6)
            records = self._read_all_records(SYSOBJ_SYSCOMMENTS)
            for rec in records:
                fixed = rec['fixed']
                if len(fixed) < 8:
                    continue
                obj_id = struct.unpack_from('<i', fixed, 0)[0]
                number = struct.unpack_from('<h', fixed, 4)[0]
                colid = struct.unpack_from('<h', fixed, 6)[0]

                text = ''
                if rec['var_data']:
                    for vd in rec['var_data']:
                        if len(vd) >= 2:
                            decoded = _smart_decode(vd)
                            if decoded and len(decoded) > len(text):
                                text = decoded

                if text:
                    self._comments[obj_id].append((number, colid, text))

        # Sort by number then colid for proper ordering
        for oid in self._comments:
            self._comments[oid].sort(key=lambda x: (x[0], x[1]))

        return self._comments

    def get_sysindexes(self):
        """Parse sysindexes (objId=2). Returns list of index dicts."""
        records = self._read_all_records(SYSOBJ_SYSINDEXES)
        indexes = []

        for rec in records:
            fixed = rec['fixed']
            if len(fixed) < 10:
                continue

            # fixed[0:4] = id (int) — the table this index belongs to.
            table_id = struct.unpack_from('<i', fixed, 0)[0]

            # indid's byte position within the fixed area shifts between
            # versions and isn't reliably decoded here; fixed_hex is exposed
            # below for callers that need to inspect the raw record.
            indid = 0

            # Index name is the first variable-length column.
            name = ''
            if rec['var_data']:
                try:
                    name = rec['var_data'][0].decode('utf-16le', errors='replace').rstrip('\x00')
                except Exception:
                    name = '<unknown>'

            indexes.append({
                'table_id': table_id,
                'name': name,
                'indid': indid,
                'fixed_hex': fixed[:20].hex(),
            })

        return indexes

    def get_table_list(self):
        """Get list of user tables: [(obj_id, name), ...]."""
        objects = self.get_objects()
        tables = []
        for oid, (name, otype) in sorted(objects.items()):
            if otype == 'U':
                tables.append((oid, name))
        return tables

    def get_proc_list(self):
        """Get list of stored procedures: [(obj_id, name), ...]."""
        objects = self.get_objects()
        procs = []
        for oid, (name, otype) in sorted(objects.items()):
            if otype == 'P':
                procs.append((oid, name))
        return procs

    def get_view_list(self):
        """Get list of views: [(obj_id, name), ...]."""
        objects = self.get_objects()
        views = []
        for oid, (name, otype) in sorted(objects.items()):
            if otype == 'V':
                views.append((oid, name))
        return views

    def get_table_id(self, table_name):
        """Look up table object ID by name (case-insensitive)."""
        objects = self.get_objects()
        name_lower = table_name.lower()
        for oid, (name, otype) in objects.items():
            if name.lower() == name_lower and otype == 'U':
                return oid
        return None

    def get_table_schema(self, table_name_or_id):
        """Get ordered column definitions for a table."""
        if isinstance(table_name_or_id, str):
            tid = self.get_table_id(table_name_or_id)
        else:
            tid = table_name_or_id

        if tid is None:
            return None

        columns = self.get_columns()
        return columns.get(tid, [])

    def get_proc_text(self, proc_name_or_id):
        """Get stored procedure SQL text."""
        objects = self.get_objects()
        comments = self.get_comments()

        if isinstance(proc_name_or_id, str):
            proc_id = None
            name_lower = proc_name_or_id.lower()
            for oid, (name, otype) in objects.items():
                if name.lower() == name_lower and otype in ('P', 'V'):
                    proc_id = oid
                    break
            if proc_id is None:
                return None
        else:
            proc_id = proc_name_or_id

        segments = comments.get(proc_id, [])
        if not segments:
            return None

        # Concatenate text segments in order
        return ''.join(text for _, _, text in segments)


class TableReader:
    """High-level table data reader."""

    def __init__(self, mdf, catalog):
        self.mdf = mdf
        self.catalog = catalog

    def _get_column_layout(self, schema):
        """Separate columns into fixed and variable, compute their positions.

        Returns (fixed_cols, var_cols) where each is a list of
        (col_def, position_info) tuples.
        """
        fixed_cols = []
        var_cols = []

        for col in schema:
            if col.xtype in VARIABLE_TYPES:
                var_cols.append(col)
            elif col.xtype in LOB_TYPES:
                var_cols.append(col)
            else:
                fixed_cols.append(col)

        # Sort fixed columns by their record offset
        fixed_cols.sort(key=lambda c: c.offset if c.offset > 0 else 9999)
        # Variable columns maintain colid order
        var_cols.sort(key=lambda c: c.colid)

        return fixed_cols, var_cols

    def read_table(self, table_name, max_rows=None):
        """Read all rows from a table. Returns list of dicts."""
        tid = self.catalog.get_table_id(table_name)
        if tid is None:
            raise ValueError(f'Table not found: {table_name}')

        schema = self.catalog.get_table_schema(tid)
        if not schema:
            raise ValueError(f'No schema found for table: {table_name}')

        fixed_cols, var_cols = self._get_column_layout(schema)
        rows = []

        for pn, hdr in self.mdf.iter_data_pages(tid):
            if hdr.slotCnt <= 0:
                continue

            page_data = self.mdf.read_page_raw(pn)
            slots = self.mdf.read_slot_array(pn, hdr.slotCnt)

            for slot_off in slots:
                if slot_off < HEADER_SIZE or slot_off >= PAGE_SIZE:
                    continue

                rec = parse_record(page_data, slot_off, hdr.pminlen)
                if rec is None:
                    continue

                row = self._decode_row(rec, schema, fixed_cols, var_cols)
                if row is not None:
                    rows.append(row)

                if max_rows and len(rows) >= max_rows:
                    return rows

        return rows

    def iter_table(self, table_name):
        """Generator version of read_table."""
        tid = self.catalog.get_table_id(table_name)
        if tid is None:
            raise ValueError(f'Table not found: {table_name}')

        schema = self.catalog.get_table_schema(tid)
        if not schema:
            raise ValueError(f'No schema found for table: {table_name}')

        fixed_cols, var_cols = self._get_column_layout(schema)

        for pn, hdr in self.mdf.iter_data_pages(tid):
            if hdr.slotCnt <= 0:
                continue

            page_data = self.mdf.read_page_raw(pn)
            slots = self.mdf.read_slot_array(pn, hdr.slotCnt)

            seen_offsets = set()
            for slot_off in slots:
                if slot_off < HEADER_SIZE or slot_off >= PAGE_SIZE:
                    continue
                if slot_off in seen_offsets:
                    continue  # skip duplicate slot pointers (MDF page splits)
                seen_offsets.add(slot_off)

                rec = parse_record(page_data, slot_off, hdr.pminlen)
                if rec is None:
                    continue

                row = self._decode_row(rec, schema, fixed_cols, var_cols)
                if row is not None:
                    yield row

    def _decode_row(self, rec, schema, fixed_cols, var_cols):
        """Decode a parsed record into a dict using the table schema."""
        row = {}
        null_bm = rec['null_bitmap']

        # Decode fixed-length columns
        for col in fixed_cols:
            col_idx = col.colid - 1  # 0-based index for null bitmap

            if is_null(null_bm, col_idx):
                row[col.name] = None
                continue

            # The offset field from syscolumns gives the byte position in the record
            # (including the 4-byte header)
            off = col.offset
            if off <= 0:
                # Negative offset usually means variable column
                # but we classified it as fixed — skip it
                row[col.name] = None
                continue

            # Determine how many bytes to read
            if col.xtype in FIXED_TYPE_SIZES:
                size = FIXED_TYPE_SIZES[col.xtype]
            elif col.xtype in FIXED_WITH_LENGTH_TYPES:
                size = col.length
            elif col.xtype in (XTYPE_DECIMAL, XTYPE_NUMERIC):
                size = col.length
            else:
                size = col.length

            if size <= 0:
                row[col.name] = None
                continue

            # Read from raw record data
            raw = rec['raw']
            if off + size > len(raw):
                row[col.name] = None
                continue

            raw_bytes = raw[off:off + size]
            row[col.name] = decode_value(col.xtype, raw_bytes, col.length)

        # Decode variable-length columns
        var_data = rec['var_data']
        for vi, col in enumerate(var_cols):
            col_idx = col.colid - 1

            if is_null(null_bm, col_idx):
                row[col.name] = None
                continue

            # Use negative offset to index into var_data if available:
            # offset=-1 → var_data[0], offset=-2 → var_data[1], etc.
            if col.offset < 0:
                var_idx = (-col.offset) - 1
            else:
                var_idx = vi  # fallback to sequential

            if var_idx < len(var_data):
                raw_bytes = var_data[var_idx]
                if col.xtype in LOB_TYPES and len(raw_bytes) <= 24:
                    # LOB pointer — data is on separate pages
                    row[col.name] = f'<LOB:{len(raw_bytes)}bytes>'
                else:
                    row[col.name] = decode_value(col.xtype, raw_bytes, col.length)
            else:
                row[col.name] = None

        return row

    def count_rows(self, table_name):
        """Count rows in a table without decoding."""
        tid = self.catalog.get_table_id(table_name)
        if tid is None:
            return 0

        count = 0
        for pn, hdr in self.mdf.iter_data_pages(tid):
            count += max(0, hdr.slotCnt)
        return count


def format_csv(rows, columns=None, file=None):
    """Write rows as CSV."""
    if not rows:
        return

    if file is None:
        file = sys.stdout

    if columns is None:
        columns = list(rows[0].keys())

    writer = csv.DictWriter(file, fieldnames=columns, extrasaction='ignore',
                             quoting=csv.QUOTE_MINIMAL, escapechar='\\')
    writer.writeheader()
    for row in rows:
        # Convert non-string values for CSV
        csv_row = {}
        for k, v in row.items():
            if isinstance(v, datetime):
                csv_row[k] = v.isoformat()
            elif isinstance(v, bool):
                csv_row[k] = '1' if v else '0'
            elif v is None:
                csv_row[k] = ''
            else:
                csv_row[k] = v
            csv_row[k] = str(csv_row[k]) if csv_row[k] != '' else ''
        writer.writerow(csv_row)


def generate_ddl(catalog, reader):
    """Generate CREATE TABLE + indexes + stored procedure DDL."""
    lines = []
    objects = catalog.get_objects()

    # CREATE TABLE statements
    for tid, tname in catalog.get_table_list():
        schema = catalog.get_table_schema(tid)
        if not schema:
            continue

        lines.append(f'CREATE TABLE [{tname}] (')
        col_lines = []
        for col in schema:
            nullable = 'NULL' if col.is_nullable else 'NOT NULL'
            col_lines.append(f'    [{col.name}] {col.type_str()} {nullable}')
        lines.append(',\n'.join(col_lines))
        lines.append(');')
        lines.append('')

    # Stored procedures
    for pid, pname in catalog.get_proc_list():
        text = catalog.get_proc_text(pid)
        if text:
            lines.append(f'-- Stored Procedure: {pname}')
            lines.append(text)
            lines.append('GO')
            lines.append('')

    # Views
    for vid, vname in catalog.get_view_list():
        text = catalog.get_proc_text(vid)
        if text:
            lines.append(f'-- View: {vname}')
            lines.append(text)
            lines.append('GO')
            lines.append('')

    return '\n'.join(lines)


def cmd_info(mdf):
    """Show database info."""
    db_name = mdf.get_db_name()
    version = mdf.get_version()

    version_names = {
        539: 'SQL Server 2000',
        611: 'SQL Server 2005',
        612: 'SQL Server 2005 SP2+',
        655: 'SQL Server 2008',
        661: 'SQL Server 2008 R2',
    }

    print(f'File:     {mdf.path}')
    print(f'Size:     {mdf.size:,} bytes ({mdf.size / 1048576:.1f} MB)')
    print(f'Pages:    {mdf.page_count:,}')
    print(f'DB Name:  {db_name}')
    print(f'Version:  {version} ({version_names.get(version, "unknown")})')

    # Page type distribution
    type_counts = defaultdict(int)
    type_names = {
        1: 'DATA', 2: 'INDEX', 3: 'LOB', 8: 'PFS', 9: 'GAM/SGAM',
        10: 'IAM', 11: 'PFS', 13: 'DCM', 15: 'BOOT', 16: 'DIFF_MAP',
        17: 'ML_MAP',
    }
    for pn in range(mdf.page_count):
        hdr = mdf.read_page_header(pn)
        if hdr:
            type_counts[hdr.type] += 1

    print(f'\nPage types:')
    for pt in sorted(type_counts):
        tname = type_names.get(pt, f'type_{pt}')
        print(f'  {tname:12s} (type {pt:2d}): {type_counts[pt]:,} pages')


def cmd_tables(mdf, catalog, reader):
    """List all user tables with row counts."""
    tables = catalog.get_table_list()
    if not tables:
        print('No user tables found.')
        return

    print(f'{"Table":<40s} {"ObjId":>10s} {"Rows":>8s} {"Columns":>8s}')
    print('-' * 70)

    for tid, tname in tables:
        schema = catalog.get_table_schema(tid)
        ncols = len(schema) if schema else 0
        nrows = reader.count_rows(tname)
        print(f'{tname:<40s} {tid:>10d} {nrows:>8d} {ncols:>8d}')


def cmd_schema(mdf, catalog, table_name):
    """Show column definitions for a table."""
    tid = catalog.get_table_id(table_name)
    if tid is None:
        print(f'Table not found: {table_name}', file=sys.stderr)
        sys.exit(1)

    schema = catalog.get_table_schema(tid)
    if not schema:
        print(f'No columns found for table: {table_name}', file=sys.stderr)
        sys.exit(1)

    objects = catalog.get_objects()
    tname = objects.get(tid, (table_name, 'U'))[0]
    print(f'Table: {tname} (objId={tid})')
    print(f'{"#":<4s} {"Column":<30s} {"Type":<20s} {"Length":>6s} {"Offset":>6s} {"Var":>3s}')
    print('-' * 73)

    for col in schema:
        var = 'Y' if col.is_variable else 'N'
        print(f'{col.colid:<4d} {col.name:<30s} {col.type_str():<20s} {col.length:>6d} {col.offset:>6d} {var:>3s}')


def cmd_dump(mdf, catalog, reader, table_name, outfile=None):
    """Dump a table as CSV."""
    schema = catalog.get_table_schema(table_name)
    if schema is None:
        print(f'Table not found: {table_name}', file=sys.stderr)
        sys.exit(1)

    columns = [col.name for col in schema]
    rows = reader.read_table(table_name)

    if outfile:
        with open(outfile, 'w', newline='', encoding='utf-8') as f:
            format_csv(rows, columns, file=f)
        print(f'Wrote {len(rows)} rows to {outfile}', file=sys.stderr)
    else:
        format_csv(rows, columns)


def cmd_dump_all(mdf, catalog, reader, outdir):
    """Dump all user tables as CSV files."""
    os.makedirs(outdir, exist_ok=True)
    tables = catalog.get_table_list()

    for tid, tname in tables:
        schema = catalog.get_table_schema(tid)
        if not schema:
            continue

        safe_name = tname.replace('/', '_').replace('\\', '_').replace('\x00', '')
        if not safe_name:
            safe_name = f'table_{tid}'
        outfile = os.path.join(outdir, f'{safe_name}.csv')
        columns = [col.name for col in schema]

        try:
            rows = reader.read_table(tname)
            with open(outfile, 'w', newline='', encoding='utf-8') as f:
                format_csv(rows, columns, file=f)
            print(f'  {tname}: {len(rows)} rows → {outfile}')
        except Exception as e:
            print(f'  {tname}: ERROR — {e}', file=sys.stderr)


def cmd_procs(mdf, catalog):
    """List stored procedures."""
    procs = catalog.get_proc_list()
    if not procs:
        print('No stored procedures found.')
        return

    comments = catalog.get_comments()
    print(f'{"Procedure":<50s} {"ObjId":>10s} {"Segments":>8s}')
    print('-' * 72)
    for pid, pname in procs:
        segs = len(comments.get(pid, []))
        print(f'{pname:<50s} {pid:>10d} {segs:>8d}')


def cmd_dump_proc(mdf, catalog, proc_name):
    """Dump a stored procedure's SQL text."""
    text = catalog.get_proc_text(proc_name)
    if text is None:
        print(f'Procedure not found: {proc_name}', file=sys.stderr)
        sys.exit(1)
    print(text)


def cmd_dump_procs(mdf, catalog, outdir):
    """Dump all stored procedures to .sql files."""
    os.makedirs(outdir, exist_ok=True)
    procs = catalog.get_proc_list()
    views = catalog.get_view_list()

    for pid, pname in procs + views:
        text = catalog.get_proc_text(pid)
        if not text:
            continue

        safe_name = pname.replace('/', '_').replace('\\', '_').replace('\x00', '')
        if not safe_name:
            safe_name = f'proc_{pid}'
        outfile = os.path.join(outdir, f'{safe_name}.sql')
        with open(outfile, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f'  {pname} → {outfile}')


def cmd_indexes(mdf, catalog):
    """List indexes."""
    indexes = catalog.get_sysindexes()
    objects = catalog.get_objects()

    if not indexes:
        print('No indexes found.')
        return

    print(f'{"Table":<30s} {"Index":<30s} {"TableId":>10s}')
    print('-' * 74)
    for idx in indexes:
        tname = objects.get(idx['table_id'], (f'obj_{idx["table_id"]}', '?'))[0]
        print(f'{tname:<30s} {idx["name"]:<30s} {idx["table_id"]:>10d}')


def cmd_foreign_keys(mdf, catalog):
    """List foreign key relationships."""
    objects = catalog.get_objects()
    # FK constraints have type 'F' in sysobjects
    fks = [(oid, name) for oid, (name, otype) in objects.items() if otype == 'F']

    if not fks:
        print('No foreign keys found.')
        return

    print(f'{"FK Name":<40s} {"ObjId":>10s}')
    print('-' * 54)
    for fkid, fkname in sorted(fks):
        print(f'{fkname:<40s} {fkid:>10d}')


def cmd_ddl(mdf, catalog, reader):
    """Generate full DDL dump."""
    print(generate_ddl(catalog, reader))


def _sqlite_type(col):
    """Map SQL Server xtype to SQLite type affinity."""
    if col.xtype in (XTYPE_TINYINT, XTYPE_SMALLINT, XTYPE_INT, XTYPE_BIGINT, XTYPE_BIT):
        return 'INTEGER'
    if col.xtype in (XTYPE_REAL, XTYPE_FLOAT, XTYPE_MONEY, XTYPE_SMALLMONEY,
                      XTYPE_DECIMAL, XTYPE_NUMERIC):
        return 'REAL'
    if col.xtype in (XTYPE_BINARY, XTYPE_VARBINARY, XTYPE_IMAGE, XTYPE_TIMESTAMP):
        return 'BLOB'
    # text types: varchar, char, nvarchar, nchar, text, ntext, datetime, etc.
    return 'TEXT'


def _safe_table_name(name):
    """Sanitize a table name for SQLite."""
    name = name.replace('\x00', '').strip()
    if not name:
        return None
    return name


def _quote_ident(name):
    """Quote a SQL identifier for SQLite using double-quote style.

    Double-quote quoting handles all characters including ] which
    bracket quoting [name] cannot escape. Double-quotes inside the
    name are escaped as "".
    """
    return '"' + name.replace('"', '""') + '"'


def cmd_sqlite(mdf, catalog, reader, outpath):
    """Export MDF to a SQLite database."""
    version = mdf.get_version()
    db_name = mdf.get_db_name()
    version_names = {
        539: 'SQL Server 2000', 611: 'SQL Server 2005',
        612: 'SQL Server 2005 SP2+', 655: 'SQL Server 2008',
        661: 'SQL Server 2008 R2',
    }

    conn = sqlite3.connect(outpath)
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA synchronous=NORMAL')

    # -- metadata table --
    conn.execute('''CREATE TABLE _metadata (
        key TEXT PRIMARY KEY,
        value TEXT
    )''')
    meta = [
        ('db_name', db_name),
        ('source_file', os.path.basename(mdf.path)),
        ('source_size', str(mdf.size)),
        ('mdf_version', str(version)),
        ('mdf_version_name', version_names.get(version, f'v{version}')),
        ('page_count', str(mdf.page_count)),
    ]
    conn.executemany('INSERT INTO _metadata VALUES (?, ?)', meta)

    # -- schema table --
    conn.execute('''CREATE TABLE _schema (
        table_name TEXT,
        column_id INTEGER,
        column_name TEXT,
        type_name TEXT,
        xtype INTEGER,
        length INTEGER,
        prec INTEGER,
        scale INTEGER,
        is_variable INTEGER,
        is_nullable INTEGER
    )''')

    tables = catalog.get_table_list()
    # Schema rows populated during table creation to capture renames
    schema_rows = []

    # -- procedures table --
    conn.execute('''CREATE TABLE _procedures (
        name TEXT,
        obj_id INTEGER,
        type TEXT,
        sql_text TEXT
    )''')
    for pid, pname in catalog.get_proc_list():
        text = catalog.get_proc_text(pid)
        if text:
            conn.execute('INSERT INTO _procedures VALUES (?,?,?,?)',
                         (pname, pid, 'P', text))
    for vid, vname in catalog.get_view_list():
        text = catalog.get_proc_text(vid)
        if text:
            conn.execute('INSERT INTO _procedures VALUES (?,?,?,?)',
                         (vname, vid, 'V', text))

    # -- user data tables --
    exported = 0
    for tid, tname in tables:
        tname_safe = _safe_table_name(tname)
        if not tname_safe:
            continue
        schema = catalog.get_table_schema(tid)
        if not schema:
            continue

        # Build CREATE TABLE with SQLite types
        col_defs = []
        col_names = []
        seen_cols = set()
        for col in schema:
            safe_col = col.name.replace('\x00', '').strip() if col.name else ''
            if not safe_col:
                safe_col = f'col_{col.colid}'
            # Handle duplicate column names (can occur from MDF corruption)
            orig = safe_col
            suffix = 2
            while safe_col in seen_cols:
                safe_col = f'{orig}_{suffix}'
                suffix += 1
            seen_cols.add(safe_col)
            col_names.append(safe_col)
            sqlite_type = _sqlite_type(col)
            col_defs.append(f'{_quote_ident(safe_col)} {sqlite_type}')

        create_sql = f'CREATE TABLE {_quote_ident(tname_safe)} ({", ".join(col_defs)})'
        try:
            conn.execute(create_sql)
        except sqlite3.OperationalError:
            # Table name conflict (e.g. duplicate after sanitization)
            tname_safe = f'{tname_safe}_{tid}'
            create_sql = f'CREATE TABLE {_quote_ident(tname_safe)} ({", ".join(col_defs)})'
            try:
                conn.execute(create_sql)
            except sqlite3.OperationalError:
                continue

        # Record schema with final (possibly renamed) table name
        for col in schema:
            schema_rows.append((
                tname_safe, col.colid, col.name, col.type_str(),
                col.xtype, col.length, col.prec, col.scale,
                int(col.is_variable), int(col.is_nullable),
            ))

        placeholders = ', '.join(['?'] * len(col_names))
        insert_sql = f'INSERT INTO {_quote_ident(tname_safe)} VALUES ({placeholders})'

        batch = []
        row_count = 0
        for row in reader.iter_table(tname):
            values = []
            for col in schema:
                v = row.get(col.name)
                if isinstance(v, datetime):
                    v = v.isoformat()
                elif isinstance(v, bytes):
                    v = sqlite3.Binary(v)
                values.append(v)
            batch.append(values)
            row_count += 1
            if len(batch) >= 1000:
                conn.executemany(insert_sql, batch)
                batch.clear()

        if batch:
            conn.executemany(insert_sql, batch)

        exported += 1
        print(f'  {tname_safe}: {row_count} rows', file=sys.stderr)

    conn.executemany('INSERT INTO _schema VALUES (?,?,?,?,?,?,?,?,?,?)', schema_rows)
    conn.commit()
    conn.close()
    print(f'Exported {exported} tables to {outpath}', file=sys.stderr)


def cmd_scan_dir(scan_dir, outdir=None, sqlite=False):
    """Scan directory for MDF files and show summary."""
    mdf_files = []
    for root, dirs, files in os.walk(scan_dir):
        for fname in files:
            if fname.lower().endswith('.mdf'):
                mdf_files.append(os.path.join(root, fname))

    if not mdf_files:
        print(f'No MDF files found in {scan_dir}')
        return

    mdf_files.sort()
    print(f'Found {len(mdf_files)} MDF files\n')

    version_names = {
        539: 'SQL2000', 611: 'SQL2005', 612: 'SQL2005SP2',
        655: 'SQL2008', 661: 'SQL2008R2',
    }

    total_size = 0
    total_tables = 0
    errors = 0

    print(f'{"File":<60s} {"Size":>10s} {"Ver":>10s} {"Tables":>6s} {"Procs":>6s}')
    print('-' * 96)

    for mdf_path in mdf_files:
        rel_path = os.path.relpath(mdf_path, scan_dir)
        size = os.path.getsize(mdf_path)
        total_size += size

        try:
            with MdfFile(mdf_path) as mdf:
                version = mdf.get_version()
                ver_str = version_names.get(version, f'v{version}')

                catalog = SystemCatalog(mdf)
                ntables = len(catalog.get_table_list())
                nprocs = len(catalog.get_proc_list())
                total_tables += ntables

                print(f'{rel_path:<60s} {size:>10,} {ver_str:>10s} {ntables:>6d} {nprocs:>6d}')

                if outdir and sqlite:
                    # Export each MDF to a SQLite database
                    reader = TableReader(mdf, catalog)
                    db_name = mdf.get_db_name()
                    safe_name = rel_path.replace('/', '_').replace('\\', '_')
                    if safe_name.lower().endswith('.mdf'):
                        safe_name = safe_name[:-4]
                    os.makedirs(outdir, exist_ok=True)
                    db_path = os.path.join(outdir, f'{safe_name}.db')
                    cmd_sqlite(mdf, catalog, reader, db_path)
                elif outdir:
                    # Export each MDF to CSV files in a subdirectory
                    safe_name = rel_path.replace('/', '_').replace('\\', '_')
                    if safe_name.lower().endswith('.mdf'):
                        safe_name = safe_name[:-4]
                    export_dir = os.path.join(outdir, safe_name)
                    os.makedirs(export_dir, exist_ok=True)

                    reader = TableReader(mdf, catalog)
                    for tid, tname in catalog.get_table_list():
                        schema = catalog.get_table_schema(tid)
                        if not schema:
                            continue
                        try:
                            rows = reader.read_table(tname)
                            if rows:
                                safe_tname = tname.replace('/', '_').replace('\\', '_')
                                csv_path = os.path.join(export_dir, f'{safe_tname}.csv')
                                columns = [col.name for col in schema]
                                with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                                    format_csv(rows, columns, file=f)
                        except Exception:
                            pass

        except Exception as e:
            errors += 1
            print(f'{rel_path:<60s} {size:>10,} {"ERROR":>10s} — {e}')

    print(f'\nSummary: {len(mdf_files)} files, {total_size / 1048576:.1f} MB, '
          f'{total_tables} tables, {errors} errors')


def cmd_diff(old_path, new_path, table=None, as_json=False):
    """Diff two SQLite database exports."""
    if not os.path.exists(old_path):
        print(f'File not found: {old_path}', file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(new_path):
        print(f'File not found: {new_path}', file=sys.stderr)
        sys.exit(1)

    old_conn = sqlite3.connect(f'file:{old_path}?mode=ro', uri=True)
    new_conn = sqlite3.connect(f'file:{new_path}?mode=ro', uri=True)
    old_conn.row_factory = sqlite3.Row
    new_conn.row_factory = sqlite3.Row

    # --- Read metadata ---
    def read_meta(conn):
        meta = {}
        try:
            for row in conn.execute('SELECT key, value FROM _metadata'):
                meta[row['key']] = row['value']
        except sqlite3.OperationalError:
            pass
        return meta

    old_meta = read_meta(old_conn)
    new_meta = read_meta(new_conn)

    # --- Read schema ---
    def read_schema(conn):
        """Returns {table_name: [(col_name, type_name), ...]}"""
        schema = {}
        try:
            for row in conn.execute('SELECT table_name, column_name, type_name FROM _schema ORDER BY table_name, column_id'):
                tname = row['table_name']
                if tname not in schema:
                    schema[tname] = []
                schema[tname].append((row['column_name'], row['type_name']))
        except sqlite3.OperationalError:
            pass
        return schema

    old_schema = read_schema(old_conn)
    new_schema = read_schema(new_conn)

    old_tables = set(old_schema.keys())
    new_tables = set(new_schema.keys())
    added_tables = sorted(new_tables - old_tables)
    removed_tables = sorted(old_tables - new_tables)
    common_tables = sorted(old_tables & new_tables)

    # --- Row counts for common tables ---
    def row_count(conn, tname):
        try:
            r = conn.execute(f'SELECT COUNT(*) FROM [{tname}]').fetchone()
            return r[0]
        except sqlite3.OperationalError:
            return 0

    # --- Per-table schema + row analysis ---
    table_changes = []
    for tname in common_tables:
        old_cols = old_schema[tname]
        new_cols = new_schema[tname]
        old_col_names = [c[0] for c in old_cols]
        new_col_names = [c[0] for c in new_cols]
        old_col_map = {c[0]: c[1] for c in old_cols}
        new_col_map = {c[0]: c[1] for c in new_cols}

        added_cols = [c for c in new_col_names if c not in old_col_map]
        removed_cols = [c for c in old_col_names if c not in new_col_map]
        type_changed = []
        for c in new_col_names:
            if c in old_col_map and old_col_map[c] != new_col_map[c]:
                type_changed.append((c, old_col_map[c], new_col_map[c]))

        old_rows = row_count(old_conn, tname)
        new_rows = row_count(new_conn, tname)

        if added_cols or removed_cols or type_changed or old_rows != new_rows:
            table_changes.append({
                'name': tname,
                'old_cols': len(old_cols), 'new_cols': len(new_cols),
                'added_cols': added_cols, 'removed_cols': removed_cols,
                'type_changed': type_changed,
                'old_rows': old_rows, 'new_rows': new_rows,
            })

    # --- Procedures ---
    def read_procs(conn):
        procs = {}
        try:
            for row in conn.execute('SELECT name, sql_text FROM _procedures'):
                procs[row['name']] = row['sql_text']
        except sqlite3.OperationalError:
            pass
        return procs

    old_procs = read_procs(old_conn)
    new_procs = read_procs(new_conn)
    old_proc_names = set(old_procs.keys())
    new_proc_names = set(new_procs.keys())
    added_procs = sorted(new_proc_names - old_proc_names)
    removed_procs = sorted(old_proc_names - new_proc_names)
    modified_procs = sorted(n for n in old_proc_names & new_proc_names
                            if old_procs[n] != new_procs[n])

    # --- Detailed row diff for --table ---
    row_diff = None
    if table:
        tname = table
        # Find case-insensitive match
        all_tables = old_tables | new_tables
        match = None
        for t in all_tables:
            if t.lower() == tname.lower():
                match = t
                break
        if match is None:
            print(f'Table not found in either database: {tname}', file=sys.stderr)
            sys.exit(1)
        tname = match

        if tname not in old_tables:
            print(f'Table {tname} only exists in new database (all rows are new).', file=sys.stderr)
            sys.exit(1)
        if tname not in new_tables:
            print(f'Table {tname} only exists in old database (all rows removed).', file=sys.stderr)
            sys.exit(1)

        # Column intersection (case-insensitive)
        old_col_names = [c[0] for c in old_schema[tname]]
        new_col_names = [c[0] for c in new_schema[tname]]
        new_col_lower = {c.lower(): c for c in new_col_names}
        # Build list of (old_name, new_name) pairs for common columns
        common_pairs = []
        for c in old_col_names:
            if c.lower() in new_col_lower:
                common_pairs.append((c, new_col_lower[c.lower()]))
        if not common_pairs:
            print(f'No common columns between old and new for table {tname}', file=sys.stderr)
            sys.exit(1)

        common_cols = [p[0] for p in common_pairs]  # use old names as canonical
        old_col_list = ', '.join(_quote_ident(p[0]) for p in common_pairs)
        new_col_list = ', '.join(_quote_ident(p[1]) for p in common_pairs)
        key_col = common_cols[0]

        def _norm_val(v):
            """Normalize a value for comparison across type changes."""
            if v is None:
                return None
            if isinstance(v, float) and v == int(v):
                return str(int(v))  # 0.0 -> '0', 10.0 -> '10'
            return str(v)

        def read_rows_keyed(conn, clist, tbl, kcol):
            rows = {}
            dupes = False
            try:
                for row in conn.execute(f'SELECT {clist} FROM {_quote_ident(tbl)}'):
                    key = row[kcol]
                    vals = {c: row[c] for c in common_cols}
                    if key in rows:
                        dupes = True
                    rows[key] = vals
            except sqlite3.OperationalError:
                pass
            return rows, dupes

        old_rows_map, old_dupes = read_rows_keyed(old_conn, old_col_list, tname, key_col)
        new_rows_map, new_dupes = read_rows_keyed(new_conn, new_col_list, tname, key_col)

        # If duplicates in key column, fall back to row-tuple multiset comparison
        if old_dupes or new_dupes:
            key_col = '(row_hash)'

            def read_row_tuples(conn, clist, tbl):
                tuples = []
                try:
                    for row in conn.execute(f'SELECT {clist} FROM {_quote_ident(tbl)}'):
                        tuples.append(tuple(_norm_val(row[c])
                                            for c in common_cols))
                except sqlite3.OperationalError:
                    pass
                return tuples

            old_tuples = read_row_tuples(old_conn, old_col_list, tname)
            new_tuples = read_row_tuples(new_conn, new_col_list, tname)
            old_set = defaultdict(int)
            new_set = defaultdict(int)
            for t in old_tuples:
                old_set[t] += 1
            for t in new_tuples:
                new_set[t] += 1

            all_tuples = set(old_set.keys()) | set(new_set.keys())
            added_keys = []
            removed_keys = []
            for t in all_tuples:
                diff = new_set.get(t, 0) - old_set.get(t, 0)
                label = f'{common_cols[0]}={t[0]}'
                if diff > 0:
                    for _ in range(diff):
                        added_keys.append(label)
                elif diff < 0:
                    for _ in range(-diff):
                        removed_keys.append(label)
            added_keys.sort()
            removed_keys.sort()

            row_diff = {
                'table': tname, 'key_col': key_col,
                'common_cols': common_cols,
                'added_keys': added_keys, 'removed_keys': removed_keys,
                'modified_rows': [],
            }
        else:
            old_keys = set(old_rows_map.keys())
            new_keys = set(new_rows_map.keys())
            added_keys = sorted(new_keys - old_keys, key=lambda k: (str(type(k)), k))
            removed_keys = sorted(old_keys - new_keys, key=lambda k: (str(type(k)), k))

            modified_rows = []
            for k in sorted(old_keys & new_keys, key=lambda k: (str(type(k)), k)):
                old_r = old_rows_map[k]
                new_r = new_rows_map[k]
                changes = {}
                for c in common_cols:
                    ov, nv = old_r[c], new_r[c]
                    if _norm_val(ov) != _norm_val(nv):
                        changes[c] = (ov, nv)
                if changes:
                    modified_rows.append((k, changes))

            row_diff = {
                'table': tname, 'key_col': key_col,
                'common_cols': common_cols,
                'added_keys': added_keys, 'removed_keys': removed_keys,
                'modified_rows': modified_rows,
            }

    # === Output ===
    if as_json:
        result = {
            'old': {'path': old_path, 'meta': old_meta},
            'new': {'path': new_path, 'meta': new_meta},
            'tables': {
                'old_count': len(old_tables), 'new_count': len(new_tables),
                'added': added_tables, 'removed': removed_tables,
                'changed': [{
                    'name': tc['name'],
                    'old_cols': tc['old_cols'], 'new_cols': tc['new_cols'],
                    'added_cols': tc['added_cols'], 'removed_cols': tc['removed_cols'],
                    'type_changed': [{'col': c, 'old': o, 'new': n} for c, o, n in tc['type_changed']],
                    'old_rows': tc['old_rows'], 'new_rows': tc['new_rows'],
                } for tc in table_changes],
            },
            'procedures': {
                'old_count': len(old_procs), 'new_count': len(new_procs),
                'added': added_procs, 'removed': removed_procs,
                'modified': modified_procs,
            },
        }
        if row_diff:
            rd = row_diff
            result['row_diff'] = {
                'table': rd['table'], 'key_col': rd['key_col'],
                'added': len(rd['added_keys']),
                'removed': len(rd['removed_keys']),
                'modified': [{
                    'key': k,
                    'changes': {c: {'old': str(o), 'new': str(n)} for c, (o, n) in ch.items()},
                } for k, ch in rd['modified_rows']],
            }
        print(json.dumps(result, indent=2, default=str))
        old_conn.close()
        new_conn.close()
        return

    # --- Human-readable output ---
    def meta_label(meta, path):
        db = meta.get('db_name', '?')
        fmt = meta.get('source_format')
        if fmt:
            ver = fmt
        else:
            ver = meta.get('mdf_version_name', meta.get('mdf_version', '?'))
        size = int(meta.get('source_size', 0))
        if size:
            size_str = f'{size / 1048576:.1f}MB'
        else:
            size_str = '?'
        return f'{os.path.basename(path)} ({db}, {ver}, {size_str})'

    print(f'--- {meta_label(old_meta, old_path)}')
    print(f'+++ {meta_label(new_meta, new_path)}')
    print()

    # Table summary
    print(f'Tables: {len(old_tables)} -> {len(new_tables)} '
          f'(+{len(added_tables)} added, -{len(removed_tables)} removed)')
    for t in added_tables:
        ncols = len(new_schema.get(t, []))
        nrows = row_count(new_conn, t)
        print(f'  + {t:<40s} {ncols} cols, {nrows} rows')
    for t in removed_tables:
        ncols = len(old_schema.get(t, []))
        nrows = row_count(old_conn, t)
        print(f'  - {t:<40s} {ncols} cols, {nrows} rows')
    for tc in table_changes:
        parts = []
        if tc['old_cols'] != tc['new_cols']:
            parts.append(f'{tc["old_cols"]} -> {tc["new_cols"]} cols')
        if tc['old_rows'] != tc['new_rows']:
            parts.append(f'{tc["old_rows"]} -> {tc["new_rows"]} rows')
        if tc['type_changed']:
            parts.append(f'{len(tc["type_changed"])} type changes')
        if parts:
            print(f'  = {tc["name"]:<40s} {", ".join(parts)}')
    print()

    # Per-table schema details (for tables with column changes)
    for tc in table_changes:
        if not tc['added_cols'] and not tc['removed_cols'] and not tc['type_changed']:
            continue
        print(f'Table: {tc["name"]}')
        print(f'  Columns: {tc["old_cols"]} -> {tc["new_cols"]}')
        for c in tc['added_cols']:
            new_col_map = {col[0]: col[1] for col in new_schema[tc['name']]}
            print(f'    + {c} ({new_col_map.get(c, "?")})')
        for c in tc['removed_cols']:
            old_col_map = {col[0]: col[1] for col in old_schema[tc['name']]}
            print(f'    - {c} ({old_col_map.get(c, "?")})')
        for c, ot, nt in tc['type_changed']:
            print(f'    ~ {c}: {ot} -> {nt}')
        delta = tc['new_rows'] - tc['old_rows']
        sign = '+' if delta >= 0 else ''
        print(f'  Rows: {tc["old_rows"]} -> {tc["new_rows"]} ({sign}{delta})')
        print()

    # Procedures
    if old_procs or new_procs:
        print(f'Procedures: {len(old_procs)} -> {len(new_procs)} '
              f'(+{len(added_procs)} added, -{len(removed_procs)} removed, '
              f'~{len(modified_procs)} modified)')
        for p in added_procs:
            print(f'  + {p}')
        for p in removed_procs:
            print(f'  - {p}')
        for p in modified_procs:
            print(f'  ~ {p}: text changed')
        print()

    # Row-level diff
    if row_diff:
        rd = row_diff
        print(f'Row diff: {rd["table"]} (key: {rd["key_col"]})')
        if rd['added_keys']:
            keys_preview = rd['added_keys'][:20]
            keys_str = ', '.join(str(k) for k in keys_preview)
            more = f' ... (+{len(rd["added_keys"]) - 20} more)' if len(rd['added_keys']) > 20 else ''
            print(f'  Added {len(rd["added_keys"])} rows: {keys_str}{more}')
        if rd['removed_keys']:
            keys_preview = rd['removed_keys'][:20]
            keys_str = ', '.join(str(k) for k in keys_preview)
            more = f' ... (+{len(rd["removed_keys"]) - 20} more)' if len(rd['removed_keys']) > 20 else ''
            print(f'  Removed {len(rd["removed_keys"])} rows: {keys_str}{more}')
        if rd['modified_rows']:
            show_count = len(rd['modified_rows']) if table else min(20, len(rd['modified_rows']))
            print(f'  Modified {len(rd["modified_rows"])} rows:')
            for k, changes in rd['modified_rows'][:show_count]:
                parts = ', '.join(f'{c} {o!r}->{n!r}' for c, (o, n) in changes.items())
                print(f'    {rd["key_col"]}={k}: {parts}')
            if not table and len(rd['modified_rows']) > 20:
                print(f'    ... (+{len(rd["modified_rows"]) - 20} more, use --table for all)')
        if not rd['added_keys'] and not rd['removed_keys'] and not rd['modified_rows']:
            print('  No row-level changes (common columns identical)')
        print()

    old_conn.close()
    new_conn.close()


def main():
    parser = argparse.ArgumentParser(
        description='SQL Server MDF file parser (no SQL Server instance required)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --info database.mdf
  %(prog)s --tables database.mdf
  %(prog)s --schema Customers database.mdf
  %(prog)s --dump Customers database.mdf > customers.csv
  %(prog)s --dump-all database.mdf --outdir ./output/
  %(prog)s --procs database.mdf
  %(prog)s --dump-proc GetCustomers database.mdf
  %(prog)s --dump-procs database.mdf --outdir ./procs/
  %(prog)s --ddl database.mdf > schema.sql
  %(prog)s --sqlite database.mdf                      # → database.db
  %(prog)s --sqlite output.db database.mdf            # → output.db
  %(prog)s --scan-dir ./databases/ --sqlite --outdir ./dbs/
  %(prog)s --diff old.db new.db
  %(prog)s --diff old.db new.db --table armor
  %(prog)s --diff old.db new.db --json
""")

    parser.add_argument('mdf', nargs='?', help='Path to MDF file')
    parser.add_argument('--info', action='store_true', help='Show database info')
    parser.add_argument('--tables', action='store_true', help='List user tables with row counts')
    parser.add_argument('--schema', metavar='TABLE', help='Show column definitions for TABLE')
    parser.add_argument('--dump', metavar='TABLE', help='Dump TABLE as CSV to stdout')
    parser.add_argument('--dump-all', action='store_true', help='Dump all tables as CSV files')
    parser.add_argument('--procs', action='store_true', help='List stored procedures')
    parser.add_argument('--dump-proc', metavar='NAME', help='Dump stored procedure SQL text')
    parser.add_argument('--dump-procs', action='store_true', help='Dump all stored procedures')
    parser.add_argument('--indexes', action='store_true', help='List indexes')
    parser.add_argument('--foreign-keys', action='store_true', help='List foreign key relationships')
    parser.add_argument('--ddl', action='store_true', help='Full DDL dump')
    parser.add_argument('--sqlite', metavar='OUTPATH', nargs='?', const='AUTO',
                        help='Export to SQLite database (default: <dbname>.db)')
    parser.add_argument('--scan-dir', metavar='DIR', help='Scan directory for MDF files')
    parser.add_argument('--outdir', metavar='DIR', help='Output directory for file exports')
    parser.add_argument('--diff', nargs=2, metavar=('OLD', 'NEW'),
                        help='Diff two SQLite database exports')
    parser.add_argument('--table', metavar='TABLE',
                        help='With --diff: row-level diff for a specific table')
    parser.add_argument('--json', action='store_true',
                        help='With --diff: output as JSON')

    args = parser.parse_args()

    # --diff mode: pure SQLite comparison, no MDF needed
    if args.diff:
        cmd_diff(args.diff[0], args.diff[1], table=args.table, as_json=args.json)
        return

    # --scan-dir mode doesn't require an MDF argument
    if args.scan_dir:
        cmd_scan_dir(args.scan_dir, args.outdir, sqlite=bool(args.sqlite))
        return

    if not args.mdf:
        parser.print_help()
        sys.exit(1)

    if not os.path.exists(args.mdf):
        print(f'File not found: {args.mdf}', file=sys.stderr)
        sys.exit(1)

    with MdfFile(args.mdf) as mdf:
        catalog = SystemCatalog(mdf)
        reader = TableReader(mdf, catalog)

        if args.info:
            cmd_info(mdf)
        elif args.tables:
            cmd_tables(mdf, catalog, reader)
        elif args.schema:
            cmd_schema(mdf, catalog, args.schema)
        elif args.dump:
            cmd_dump(mdf, catalog, reader, args.dump, args.outdir)
        elif args.dump_all:
            if not args.outdir:
                args.outdir = '.'
            cmd_dump_all(mdf, catalog, reader, args.outdir)
        elif args.procs:
            cmd_procs(mdf, catalog)
        elif args.dump_proc:
            cmd_dump_proc(mdf, catalog, args.dump_proc)
        elif args.dump_procs:
            if not args.outdir:
                args.outdir = '.'
            cmd_dump_procs(mdf, catalog, args.outdir)
        elif args.indexes:
            cmd_indexes(mdf, catalog)
        elif args.foreign_keys:
            cmd_foreign_keys(mdf, catalog)
        elif args.ddl:
            cmd_ddl(mdf, catalog, reader)
        elif args.sqlite:
            if args.sqlite == 'AUTO':
                db_name = mdf.get_db_name()
                outpath = os.path.join(args.outdir or '.', f'{db_name}.db')
            else:
                outpath = args.sqlite
            cmd_sqlite(mdf, catalog, reader, outpath)
        else:
            # Default: show info + tables
            cmd_info(mdf)
            print()
            cmd_tables(mdf, catalog, reader)


if __name__ == '__main__':
    import signal
    if hasattr(signal, 'SIGPIPE'):
        signal.signal(signal.SIGPIPE, signal.SIG_DFL)
    main()
