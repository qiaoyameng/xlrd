"""
Tests for the read-only OLE container diagnostic entry point.

The fixtures build a small but fully legal OLE2 compound document from
scratch (no binary sample checked in), then apply single-point byte
mutations so that each failure mode is triggered deterministically.
"""
from __future__ import print_function

from struct import pack

import pytest

import xlrd
from xlrd import diagnose_workbook_container
from xlrd.compdoc import EOCSID, FREESID


# A well-formed BIFF8 BOF record (opcode 0x0809, 8 payload bytes) followed
# by an EOF record.  Enough for the diagnostic's BOF check; full parsing is
# deliberately out of scope here.
GOOD_BIFF = (
    pack('<HH', 0x0809, 0x0008)
    + pack('<HHHH', 0x0600, 0x0005, 0x0000, 0x0000)
    + pack('<HH', 0x000A, 0x0000)
)


class MiniOleBuilder(object):
    """
    Build a legal 512-byte-sector OLE2 file containing a Workbook stream and
    an auxiliary stream.  ``use_mini`` puts Workbook in the mini stream
    (below the 4096-byte standard-stream threshold).

    Layout (every sector id is exposed on the instance so tests can target
    exact bytes):
      0  SAT
      1  directory
      2  SSAT (mini variant only)
      3  SSCS / root mini stream (mini variant only)
      n  user stream sectors
    """

    SEC = 512
    SHORT = 64

    def __init__(self, workbook_data, extra_data, use_mini=False):
        self.use_mini = use_mini
        self.workbook_data = workbook_data
        self.extra_data = extra_data
        self._build()

    # -- construction -----------------------------------------------------
    def _build(self):
        if self.use_mini:
            self._build_mini()
        else:
            self._build_standard()

    def _build_standard(self):
        wb_sectors = self._sector_count(len(self.workbook_data))
        ex_sectors = self._sector_count(len(self.extra_data))
        # ids: 0 SAT, 1 directory, then workbook then extra
        self.sat_sid = 0
        self.dir_sid = 1
        self.workbook_first = 2
        self.extra_first = 2 + wb_sectors
        n_sectors = 2 + wb_sectors + ex_sectors

        fat = [FREESID] * n_sectors
        fat[self.sat_sid] = EOCSID
        fat[self.dir_sid] = EOCSID
        for index in range(wb_sectors):
            sid = self.workbook_first + index
            fat[sid] = (
                sid + 1 if index < wb_sectors - 1 else EOCSID
            )
        for index in range(ex_sectors):
            sid = self.extra_first + index
            fat[sid] = (
                sid + 1 if index < ex_sectors - 1 else EOCSID
            )
        self.fat = fat

        directory = self._directory(
            root_first=EOCSID, root_size=0,
            wb_first=self.workbook_first, wb_size=len(self.workbook_data),
            wb_mini=False,
            ex_first=self.extra_first, ex_size=len(self.extra_data),
            ex_mini=False,
        )
        parts = [
            self._header(
                n_sectors, dir_sid=self.dir_sid,
                ssat_first=EOCSID, ssat_count=0,
                msat_first=EOCSID, msat_count=0,
            ),
            self._fat_sector(fat),
            directory,
            self._pad(self.workbook_data),
            self._pad(self.extra_data),
        ]
        self.bytes_out = b''.join(parts)
        self._record_offsets()

    def _build_mini(self):
        wb_mini = self._mini_count(len(self.workbook_data))
        ex_mini = self._mini_count(len(self.extra_data))
        # mini sector allocation inside the SSCS
        self.workbook_first = 0
        self.extra_first = wb_mini
        sscs_payload = (
            self._pad_short(self.workbook_data)
            + self._pad_short(self.extra_data)
        )
        sscs_sectors = self._sector_count(len(sscs_payload))

        # ids: 0 SAT, 1 directory, 2 SSAT, 3.. SSCS sectors
        self.sat_sid = 0
        self.dir_sid = 1
        self.ssat_sid = 2
        self.sscs_first = 3
        n_sectors = 3 + sscs_sectors

        fat = [FREESID] * n_sectors
        fat[self.sat_sid] = EOCSID
        fat[self.dir_sid] = EOCSID
        fat[self.ssat_sid] = EOCSID
        for index in range(sscs_sectors):
            sid = self.sscs_first + index
            fat[sid] = (
                sid + 1 if index < sscs_sectors - 1 else EOCSID
            )
        self.fat = fat

        mini_fat = [FREESID] * (sscs_sectors * (self.SEC // self.SHORT))
        for index in range(wb_mini):
            sid = self.workbook_first + index
            mini_fat[sid] = (
                sid + 1 if index < wb_mini - 1 else EOCSID
            )
        for index in range(ex_mini):
            sid = self.extra_first + index
            mini_fat[sid] = (
                sid + 1 if index < ex_mini - 1 else EOCSID
            )
        self.mini_fat = mini_fat

        directory = self._directory(
            root_first=self.sscs_first, root_size=len(sscs_payload),
            wb_first=self.workbook_first, wb_size=len(self.workbook_data),
            wb_mini=True,
            ex_first=self.extra_first, ex_size=len(self.extra_data),
            ex_mini=True,
        )
        parts = [
            self._header(
                n_sectors, dir_sid=self.dir_sid,
                ssat_first=self.ssat_sid, ssat_count=1,
                msat_first=EOCSID, msat_count=0,
            ),
            self._fat_sector(fat),
            directory,
            self._int_sector(mini_fat),
            self._pad(sscs_payload),
        ]
        self.bytes_out = b''.join(parts)
        self._record_offsets()
        self.sscs_payload = sscs_payload

    # -- helpers ----------------------------------------------------------
    def _sector_count(self, size):
        return max(1, (size + self.SEC - 1) // self.SEC)

    def _mini_count(self, size):
        return max(1, (size + self.SHORT - 1) // self.SHORT)

    def _pad(self, data):
        remainder = len(data) % self.SEC
        if remainder:
            data = data + b'\x00' * (self.SEC - remainder)
        return data

    def _pad_short(self, data):
        remainder = len(data) % self.SHORT
        if remainder:
            data = data + b'\x00' * (self.SHORT - remainder)
        return data

    def _header(self, n_sectors, dir_sid, ssat_first, ssat_count,
                msat_first, msat_count):
        header = bytearray(512)
        header[0:8] = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'
        header[24:26] = pack('<H', 0x003E)   # revision
        header[26:28] = pack('<H', 0x0003)   # version
        header[28:30] = b'\xFE\xFF'          # little endian
        header[30:32] = pack('<H', 9)        # 512 byte sectors
        header[32:34] = pack('<H', 6)        # 64 byte short sectors
        header[44:48] = pack('<i', 1)        # SAT sector count
        header[48:52] = pack('<i', dir_sid)
        header[56:60] = pack('<i', 4096)     # mini stream threshold
        header[60:64] = pack('<i', ssat_first)
        header[64:68] = pack('<i', ssat_count)
        header[68:72] = pack('<i', msat_first)
        header[72:76] = pack('<i', msat_count)
        # first 109 MSAT slots: slot 0 points at the single SAT sector.
        header[76:80] = pack('<i', 0)
        for slot in range(1, 109):
            header[76 + slot * 4:80 + slot * 4] = pack('<i', FREESID)
        return bytes(header)

    def _fat_sector(self, fat):
        return self._int_sector(fat)

    def _int_sector(self, values):
        sector = bytearray()
        for value in values[:self.SEC // 4]:
            sector += pack('<i', value)
        sector += b'\x00' * (self.SEC - len(sector))
        return bytes(sector)

    def _dir_entry(self, name, etype, colour, left, right, root,
                   first_sid, size):
        encoded = name.encode('utf_16_le') + b'\x00\x00'
        entry = bytearray(128)
        entry[0:len(encoded)] = encoded
        entry[64:66] = pack('<H', len(encoded))
        entry[66] = etype
        entry[67] = colour
        entry[68:72] = pack('<i', left)
        entry[72:76] = pack('<i', right)
        entry[76:80] = pack('<i', root)
        entry[116:120] = pack('<i', first_sid)
        entry[120:124] = pack('<i', size)
        return bytes(entry)

    def _directory(self, root_first, root_size, wb_first, wb_size, wb_mini,
                   ex_first, ex_size, ex_mini):
        root = self._dir_entry(
            'Root Entry', 5, 1, -1, -1, 1, root_first, root_size
        )
        workbook = self._dir_entry(
            'Workbook', 2, 1, -1, 2, -1, wb_first, wb_size
        )
        extra = self._dir_entry(
            'ExtraData', 2, 1, -1, -1, -1, ex_first, ex_size
        )
        unused = self._dir_entry('', 0, 0, -1, -1, -1, EOCSID, 0)
        sector = root + workbook + extra + unused
        return sector + b'\x00' * (self.SEC - len(sector))

    def _record_offsets(self):
        self.sat_offset = 512 + self.sat_sid * self.SEC
        self.dir_offset = 512 + self.dir_sid * self.SEC

    def fat_entry_offset(self, sid):
        return self.sat_offset + sid * 4

    def dir_field_offset(self, did, field):
        # field: 'first' (116) or 'size' (120)
        base = self.dir_offset + did * 128
        return base + (116 if field == 'first' else 120)

    def mutate(self, offset, value):
        data = bytearray(self.bytes_out)
        data[offset] = value
        return bytes(data)

    def set_sid(self, offset, sid):
        """Overwrite a full 4-byte signed sector id (FAT/miniFAT/dir)."""
        data = bytearray(self.bytes_out)
        data[offset:offset + 4] = pack('<i', sid)
        return bytes(data)

    def truncate(self, byte_count):
        return self.bytes_out[:byte_count]


# --------------------------------------------------------------------------
# Fixtures
# --------------------------------------------------------------------------
@pytest.fixture
def standard_container():
    # Both streams are at/above the 4096-byte threshold, so they use the
    # standard FAT rather than the mini stream. Workbook spans two sectors.
    workbook = GOOD_BIFF + b'\x00' * (4608 - len(GOOD_BIFF))
    extra = b'EX' + b'\x00' * (4096 - 2)
    return MiniOleBuilder(workbook, extra, use_mini=False)


@pytest.fixture
def mini_container():
    # Both streams live below the 4096-byte threshold -> mini stream.
    workbook = GOOD_BIFF + b'\x00' * (128 - len(GOOD_BIFF))
    extra = b'XY' + b'\x00' * 62
    return MiniOleBuilder(workbook, extra, use_mini=True)


# --------------------------------------------------------------------------
# Baseline: the synthetic containers are genuinely legal OLE.
# --------------------------------------------------------------------------
class TestBaseline(object):

    def test_standard_container_is_legal_ole_and_biff(self, standard_container):
        diagnosis = diagnose_workbook_container(
            file_contents=standard_container.bytes_out
        )
        assert diagnosis.issues == []
        assert diagnosis.can_locate_workbook is True
        assert diagnosis.workbook_name == '/Workbook'
        assert diagnosis.workbook_stream_type == 'standard'
        assert diagnosis.biff_state == 'biff_ok'

    def test_mini_container_is_legal_ole_and_biff(self, mini_container):
        diagnosis = diagnose_workbook_container(
            file_contents=mini_container.bytes_out
        )
        assert diagnosis.issues == []
        assert diagnosis.can_locate_workbook is True
        assert diagnosis.workbook_stream_type == 'mini'
        assert diagnosis.biff_state == 'biff_ok'

    def test_open_workbook_accepts_synthetic_container(self, standard_container):
        # The diagnostic does not weaken parsing: xlrd itself opens the file.
        book = xlrd.open_workbook(
            file_contents=standard_container.bytes_out
        )
        assert book.nsheets == 0

    def test_filename_entry_point(self, standard_container, tmp_path):
        path = tmp_path / 'synthetic.xls'
        path.write_bytes(standard_container.bytes_out)
        diagnosis = diagnose_workbook_container(filename=str(path))
        assert diagnosis.can_locate_workbook is True

    def test_as_dict_is_json_friendly(self, standard_container):
        diagnosis = diagnose_workbook_container(
            file_contents=standard_container.bytes_out
        )
        payload = diagnosis.as_dict()
        assert payload['can_locate_workbook'] is True
        assert payload['issues'] == []
        names = [entry[1] for entry in payload['directory_entries']]
        assert 'Workbook' in names and 'ExtraData' in names


# --------------------------------------------------------------------------
# Main (standard sector) FAT chain damage
# --------------------------------------------------------------------------
class TestStandardChainDamage(object):

    def test_fat_cycle_in_workbook_chain(self, standard_container):
        # Point the workbook's second FAT entry back at its first sector.
        first = standard_container.workbook_first
        corrupted = standard_container.set_sid(
            standard_container.fat_entry_offset(first + 1), first,
        )
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        codes = [issue.code for issue in diagnosis.issues]
        assert 'stream.chain_cycle' in codes
        assert diagnosis.can_locate_workbook is False
        assert diagnosis.biff_state == 'not_checked'
        assert 'BIFF stream was never reached' in diagnosis.reason or \
            'Container is damaged' in diagnosis.reason

    def test_fat_out_of_bounds_in_workbook_chain(self, standard_container):
        first = standard_container.workbook_first
        corrupted = standard_container.set_sid(
            standard_container.fat_entry_offset(first), 9999,
        )
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        codes = [issue.code for issue in diagnosis.issues]
        assert 'stream.chain_out_of_bounds' in codes
        assert diagnosis.can_locate_workbook is False
        assert diagnosis.biff_state == 'not_checked'

    def test_fat_bad_terminator_reproduces_bare_assertion(self,
                                                          standard_container):
        # The legacy parser does `assert s == EOCSID`; a chain ending on
        # FREESID previously surfaced as a bare AssertionError.
        last = standard_container.workbook_first + 1
        corrupted = standard_container.set_sid(
            standard_container.fat_entry_offset(last), FREESID,
        )
        # Legacy contract: open_workbook still raises (AssertionError here).
        with pytest.raises(Exception):
            xlrd.open_workbook(file_contents=corrupted)
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        codes = [issue.code for issue in diagnosis.issues]
        assert 'stream.bad_terminator' in codes
        assert diagnosis.can_locate_workbook is False

    def test_seen_sector_cross_link_is_diagnosed(self, standard_container):
        # Make the extra stream's FAT entry point at a workbook sector,
        # recreating the classic "seen[sector] != 0" failure.
        extra_first = standard_container.extra_first
        corrupted = standard_container.set_sid(
            standard_container.fat_entry_offset(extra_first),
            standard_container.workbook_first,
        )
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        cross = [i for i in diagnosis.issues if i.code == 'stream.cross_link']
        assert cross
        assert cross[0].component == '/ExtraData'
        assert extra_first in cross[0].path
        assert cross[0].sector == standard_container.workbook_first

    def test_declared_size_mismatch(self, standard_container):
        # Inflate the workbook's declared size without touching the chain.
        offset = standard_container.dir_field_offset(1, 'size')
        corrupted = standard_container.mutate(offset, 0xFF)
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        codes = [issue.code for issue in diagnosis.issues]
        assert 'stream.size_mismatch' in codes
        size_issue = next(
            issue for issue in diagnosis.issues
            if issue.code == 'stream.size_mismatch'
        )
        assert size_issue.component == '/Workbook'

    def test_truncated_file_keeps_diagnosis_structured(self,
                                                        standard_container):
        # Cut the file so the workbook's second sector is missing.
        cut = 512 + standard_container.workbook_first * 512 + 512
        corrupted = standard_container.truncate(cut)
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        assert diagnosis.has_errors
        assert diagnosis.can_locate_workbook is False
        codes = [issue.code for issue in diagnosis.issues]
        assert 'sat.sector_out_of_bounds' in codes or \
            'stream.chain_out_of_bounds' in codes or \
            'header.file_truncated' in codes


# --------------------------------------------------------------------------
# Mini stream / miniFAT damage
# --------------------------------------------------------------------------
class TestMiniChainDamage(object):

    def test_mini_fat_cycle(self, mini_container):
        # miniFAT lives in the SSAT sector; redirect workbook's second short
        # sector back onto its first.
        ssat_offset = 512 + mini_container.ssat_sid * 512
        first = mini_container.workbook_first
        corrupted = mini_container.set_sid(
            ssat_offset + (first + 1) * 4, first,
        )
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        codes = [issue.code for issue in diagnosis.issues]
        assert 'mini.chain_cycle' in codes
        cycle = next(i for i in diagnosis.issues
                     if i.code == 'mini.chain_cycle')
        assert cycle.component == '/Workbook'
        assert diagnosis.can_locate_workbook is False
        assert diagnosis.biff_state == 'not_checked'

    def test_mini_fat_out_of_bounds(self, mini_container):
        ssat_offset = 512 + mini_container.ssat_sid * 512
        corrupted = mini_container.set_sid(
            ssat_offset + mini_container.workbook_first * 4, 9999,
        )
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        codes = [issue.code for issue in diagnosis.issues]
        assert 'mini.chain_out_of_bounds' in codes
        assert diagnosis.can_locate_workbook is False

    def test_mini_cross_link_between_streams(self, mini_container):
        # ExtraData's mini chain is redirected onto a workbook mini sector.
        ssat_offset = 512 + mini_container.ssat_sid * 512
        corrupted = mini_container.set_sid(
            ssat_offset + mini_container.extra_first * 4,
            mini_container.workbook_first,
        )
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        cross = [i for i in diagnosis.issues if i.code == 'mini.cross_link']
        assert cross
        assert cross[0].component == '/ExtraData'
        assert cross[0].sector == mini_container.workbook_first

    def test_mini_declared_size_mismatch(self, mini_container):
        offset = mini_container.dir_field_offset(1, 'size')
        corrupted = mini_container.mutate(offset, 0xFF)
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        codes = [issue.code for issue in diagnosis.issues]
        assert 'mini.size_mismatch' in codes


# --------------------------------------------------------------------------
# Damage in streams the Workbook does not use
# --------------------------------------------------------------------------
class TestUnrelatedStreamDamage(object):

    def test_corrupt_extra_stream_is_listed_but_workbook_still_readable(
        self, standard_container
    ):
        extra_first = standard_container.extra_first
        # Break only the auxiliary stream's chain (point outside the file).
        corrupted = standard_container.set_sid(
            standard_container.fat_entry_offset(extra_first), 9999,
        )
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        related = [
            issue for issue in diagnosis.issues
            if issue.component == '/ExtraData'
        ]
        assert any(i.code == 'stream.chain_out_of_bounds' for i in related)
        # Workbook itself is untouched and still locatable.
        assert diagnosis.can_locate_workbook is True
        assert diagnosis.biff_state == 'biff_ok'


# --------------------------------------------------------------------------
# Sound container, invalid BIFF content
# --------------------------------------------------------------------------
class TestBiffContent(object):

    def test_zero_size_workbook_is_not_claimed_readable(self):
        builder = MiniOleBuilder(
            b'', b'EX' + b'\x00' * (4096 - 2), use_mini=False
        )
        diagnosis = diagnose_workbook_container(
            file_contents=builder.bytes_out
        )
        assert diagnosis.can_locate_workbook is False
        assert diagnosis.biff_state == 'not_checked'
        assert 'no sector chain' in diagnosis.reason

    def test_sound_container_with_bad_biff(self, standard_container):
        # Corrupt only the first byte of the Workbook payload (inside its
        # first data sector), leaving every OLE structure intact.
        wb_data_offset = (
            512 + standard_container.workbook_first * 512
        )
        corrupted = standard_container.mutate(wb_data_offset, 0x00)
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        assert diagnosis.issues == []
        assert diagnosis.can_locate_workbook is True
        assert diagnosis.biff_state == 'biff_invalid'
        assert 'BIFF' in diagnosis.biff_detail or 'opcode' in \
            diagnosis.biff_detail

    def test_sound_mini_container_with_bad_biff(self, mini_container):
        # First byte of the SSCS is the first byte of the mini Workbook.
        sscs_offset = 512 + mini_container.sscs_first * 512
        corrupted = mini_container.mutate(sscs_offset, 0x00)
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        assert diagnosis.can_locate_workbook is True
        assert diagnosis.biff_state == 'biff_invalid'

    def test_legacy_parser_still_raises_on_bad_biff(self, standard_container):
        wb_data_offset = (
            512 + standard_container.workbook_first * 512
        )
        corrupted = standard_container.mutate(wb_data_offset, 0x00)
        with pytest.raises(xlrd.XLRDError):
            xlrd.open_workbook(file_contents=corrupted)


# --------------------------------------------------------------------------
# Budget and malformed-input safety
# --------------------------------------------------------------------------
class TestSafety(object):

    def test_budget_halts_cyclic_directory_without_infinite_loop(
        self, standard_container
    ):
        # Forge a directory red-black tree that points at itself: the root's
        # root_DID is already DID 1; make DID 1's left child DID 1 too.
        # DID 1 left_DID field lives at directory offset 68.
        offset = standard_container.dir_offset + 1 * 128 + 68
        corrupted = standard_container.set_sid(offset, 0x01)
        diagnosis = diagnose_workbook_container(
            file_contents=corrupted, traversal_budget=1000
        )
        codes = [issue.code for issue in diagnosis.issues]
        assert 'directory.tree_cycle' in codes
        # The corrupt red-black linkage is reported; the bounded walk returns
        # promptly instead of looping forever.
        assert diagnosis.traversal_budget_exceeded is False
        assert any(
            entry[1] == 'Workbook'
            for entry in diagnosis.directory_entries
        )

    def test_budget_exhaustion_is_reported(self, standard_container):
        diagnosis = diagnose_workbook_container(
            file_contents=standard_container.bytes_out,
            traversal_budget=1,
        )
        assert diagnosis.traversal_budget_exceeded is True
        assert any(
            issue.code.endswith('budget_exceeded') for issue in diagnosis.issues
        )

    def test_non_ole_input_returns_structured_result(self):
        diagnosis = diagnose_workbook_container(
            file_contents=b'not an ole file at all'
        )
        codes = [issue.code for issue in diagnosis.issues]
        assert 'header.bad_signature' in codes or \
            'header.header_truncated' in codes
        assert diagnosis.can_locate_workbook is False
        assert diagnosis.biff_state == 'not_checked'

    def test_raw_biff_without_ole_container_is_reported_not_parsed(self):
        # A raw BIFF record stream (no OLE signature) is a legitimate xlrd
        # input via the legacy path, but the OLE diagnostic must not pretend
        # the container is sound.
        from .helpers import from_sample
        diagnosis = diagnose_workbook_container(
            filename=from_sample('biff4_no_format_no_window2.xls')
        )
        codes = [issue.code for issue in diagnosis.issues]
        assert 'header.bad_signature' in codes
        assert diagnosis.can_locate_workbook is False
        assert diagnosis.biff_state == 'not_checked'

    def test_invalid_budget_rejected(self, standard_container):
        with pytest.raises(ValueError):
            diagnose_workbook_container(
                file_contents=standard_container.bytes_out,
                traversal_budget=0,
            )

    def test_issue_carries_sector_path(self, standard_container):
        first = standard_container.workbook_first
        corrupted = standard_container.set_sid(
            standard_container.fat_entry_offset(first), 9999,
        )
        diagnosis = diagnose_workbook_container(file_contents=corrupted)
        oob = next(
            issue for issue in diagnosis.issues
            if issue.code == 'stream.chain_out_of_bounds'
        )
        assert oob.path == [first]
        assert oob.component == '/Workbook'
