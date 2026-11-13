# -*- coding: utf-8 -*-
# Copyright (c) 2005-2012 Stephen John Machin, Lingfo Pty Ltd
# This module is part of the xlrd package, which is released under a
# BSD-style licence.
"""
Read-only diagnostics for OLE2 compound documents.

This module never repairs, rewrites or ignores damage: it inspects the OLE
container of an XLS file and returns a structured list of the problems it
finds, together with an explicit decision about whether the ``Workbook``
(or legacy ``Book``) stream can still be located and handed to the BIFF
parser.
"""
from __future__ import print_function

import os
from struct import unpack

from .compdoc import SIGNATURE, EOCSID, FREESID, MSATSID


#: Default upper bound on the number of chain-walk steps any single
#: diagnosis is allowed to take.  The budget covers MSAT, SAT, directory,
#: SSAT, SSCS and every user stream, so a corrupt cyclic directory can
#: never send the analyser into an infinite loop.
DEFAULT_TRAVERSAL_BUDGET = 10_000_000

class Issue(object):
    """A single structural problem found while diagnosing a container."""

    __slots__ = (
        'code', 'severity', 'message', 'component', 'path', 'sector',
    )

    def __init__(self, code, severity, message, component=None, path=None,
                 sector=None):
        #: Stable machine-readable identifier, see the module docstring of
        #: :mod:`xlrd.diagnostics` for the catalogue of codes.
        self.code = code
        #: ``"error"`` (reading must stop) or ``"warning"`` (reported only).
        self.severity = severity
        #: Human readable explanation.
        self.message = message
        #: Name of the directory entry / OLE structure involved, if any.
        self.component = component
        #: Sector path (a list of sector ids, or a textual path) if known.
        self.path = path
        #: Specific sector id the problem is anchored on, when applicable.
        self.sector = sector

    def __repr__(self):
        return "Issue(%s, %s, %r, component=%r, path=%r, sector=%r)" % (
            self.code, self.severity, self.message, self.component,
            self.path, self.sector,
        )

    def __eq__(self, other):
        if not isinstance(other, Issue):
            return NotImplemented
        return (
            self.code == other.code and self.severity == other.severity and
            self.message == other.message and
            self.component == other.component and self.path == other.path and
            self.sector == other.sector
        )

    def __ne__(self, other):
        result = self.__eq__(other)
        if result is NotImplemented:
            return result
        return not result

    def as_dict(self):
        """Return the issue as a plain dict suitable for JSON dumping."""
        return {
            'code': self.code,
            'severity': self.severity,
            'message': self.message,
            'component': self.component,
            'path': self.path,
            'sector': self.sector,
        }


class ContainerDiagnosis(object):
    """
    Result of :func:`diagnose_workbook_container`.

    The object is deliberately independent of the live :class:`~xlrd.compdoc.
    CompDoc` parser: producing a diagnosis never mutates the input bytes and
    never opens the file in a mode that could mask an error.
    """

    __slots__ = (
        'issues', 'can_locate_workbook', 'reason', 'workbook_name',
        'workbook_stream_type', 'workbook_path', 'biff_state',
        'biff_detail', 'directory_entries', 'traversal_budget_exceeded',
    )

    def __init__(self):
        #: list of :class:`Issue` instances, in detection order.
        self.issues = []
        #: ``True`` only when the Workbook/Book chain can be located without
        #: crossing corrupt or out-of-bounds sectors.
        self.can_locate_workbook = False
        #: Human readable justification for :attr:`can_locate_workbook`.
        self.reason = ''
        #: Actual directory name used (``"Workbook"`` or ``"Book"``) or None.
        self.workbook_name = None
        #: ``"standard"``, ``"mini"`` or ``None`` when not located.
        self.workbook_stream_type = None
        #: Sector path of the located workbook stream.
        self.workbook_path = None
        #: One of ``"biff_ok"``, ``"biff_invalid"`` or ``"not_checked"``.
        self.biff_state = 'not_checked'
        #: Explanation when :attr:`biff_state` is not ``"biff_ok"``.
        self.biff_detail = ''
        #: list of ``(DID, name, entry_type, first_sid, declared_size, path)``.
        self.directory_entries = []
        #: ``True`` if the traversal budget had to stop a chain walk.
        self.traversal_budget_exceeded = False

    @property
    def blocking_issues(self):
        """Subset of :attr:`issues` whose severity is ``"error"``."""
        return [issue for issue in self.issues if issue.severity == 'error']

    @property
    def has_errors(self):
        return bool(self.blocking_issues)

    def as_dict(self):
        """Return the whole diagnosis as a plain dict suitable for JSON."""
        return {
            'issues': [issue.as_dict() for issue in self.issues],
            'can_locate_workbook': self.can_locate_workbook,
            'reason': self.reason,
            'workbook_name': self.workbook_name,
            'workbook_stream_type': self.workbook_stream_type,
            'workbook_path': self.workbook_path,
            'biff_state': self.biff_state,
            'biff_detail': self.biff_detail,
            'directory_entries': [
                list(entry) for entry in self.directory_entries
            ],
            'traversal_budget_exceeded': self.traversal_budget_exceeded,
        }


# Issues in this set make the SAT itself untrustworthy, so chain walks that
# rely on it cannot prove a stream is safely locatable.
_FATAL_CODES = frozenset((
    'header.bad_signature',
    'header.bad_byte_order',
    'header.header_truncated',
    'header.bad_sector_size',
    'msat.out_of_bounds',
    'msat.cycle',
    'msat.bad_sid',
    'msat.budget_exceeded',
    'sat.sector_out_of_bounds',
    'sat.cycle',
    'sat.bad_sid',
    'sat.budget_exceeded',
    'sat.truncated',
    'directory.chain_out_of_bounds',
    'directory.chain_cycle',
    'directory.bad_terminator',
    'directory.budget_exceeded',
    'directory.truncated',
    'ssat.chain_out_of_bounds',
    'ssat.chain_cycle',
    'ssat.bad_terminator',
    'ssat.budget_exceeded',
    'ssat.size_mismatch',
    'sscs.chain_out_of_bounds',
    'sscs.chain_cycle',
    'sscs.bad_terminator',
    'sscs.budget_exceeded',
    'sscs.size_mismatch',
))


class _Analyzer(object):
    """Defensive, allocation-light, read-only OLE2 container analyser."""

    def __init__(self, mem, budget):
        self.mem = mem
        self.budget = budget
        self.steps = 0
        self.result = ContainerDiagnosis()
        # owner maps: sector id -> (owner label, sequence number within chain)
        self.owners = {}
        self.mini_owners = {}
        # populated as the header/directory are parsed
        self.sec_size = 512
        self.short_sec_size = 64
        self.mem_data_secs = 0
        self.min_size_std_stream = 4096
        self.dir_first_sec_sid = EOCSID
        self.ssat_first_sec_sid = EOCSID
        self.ssat_tot_secs = 0
        self.msatx_first_sec_sid = EOCSID
        self.msatx_tot_secs = 0
        self.SAT = []
        self.SSAT = []
        self.SSCS = b''
        self.dirlist = []

    def add(self, code, severity, message, component=None, path=None,
            sector=None):
        self.result.issues.append(
            Issue(
                code, severity, message, component=component, path=path,
                sector=sector,
            )
        )

    def tick(self):
        self.steps += 1
        if self.steps > self.budget:
            self.result.traversal_budget_exceeded = True
            return False
        return True

    def _read_sector(self, sid):
        start = 512 + sid * self.sec_size
        return self.mem[start:start + self.sec_size]

    def _claim(self, owners, sector_id, label, code, component, path):
        """Record ownership of a sector, flagging any cross-occupancy."""
        previous = owners.get(sector_id)
        if previous is not None:
            prev_label, prev_index = previous
            if prev_label != label:
                self.add(
                    code, 'error',
                    'Sector %d is claimed by both %r (step %d) and %r '
                    '(step %d)' % (
                        sector_id, prev_label, prev_index, label, len(path),
                    ),
                    component=component,
                    path=list(path), sector=sector_id,
                )
            else:
                self.add(
                    code, 'error',
                    'Sector %d is visited more than once by %r'
                    % (sector_id, label),
                    component=component,
                    path=list(path), sector=sector_id,
                )
            return False
        owners[sector_id] = (label, len(path))
        return True

    def _walk_fat_chain(self, start_sid, sat, n_sectors, label, owner_map,
                        claim_code, component, oob_code, cycle_code,
                        terminator_code, expect_terminator=True,
                        budget_code=None):
        """
        Walk a FAT or miniFAT chain with hard bounds.

        Returns the list of sector ids visited.  Every anomaly is recorded as
        an :class:`Issue`; the walk always terminates (cycles and the step
        budget guarantee that).
        """
        path = []
        sid = start_sid
        if budget_code is None:
            budget_code = oob_code.rsplit('.', 1)[0] + '.budget_exceeded'
        while sid >= 0:
            if not self.tick():
                self.add(
                    budget_code, 'error',
                    'Traversal budget exceeded while walking %r at sector %d'
                    % (label, sid),
                    component=component, path=list(path),
                )
                return path
            if sid >= n_sectors:
                self.add(
                    oob_code, 'error',
                    '%s references sector %d but only %d sectors exist'
                    % (label, sid, n_sectors),
                    component=component, path=list(path),
                )
                return path
            if sid in path:
                self.add(
                    cycle_code, 'error',
                    '%s contains a cycle at sector %d' % (label, sid),
                    component=component, path=list(path),
                )
                return path
            if not self._claim(
                owner_map, sid, label, claim_code, component, path
            ):
                # cross-occupancy: keep walking but this chain is already
                # proven corrupt, so stop to avoid a runaway loop.
                return path
            path.append(sid)
            try:
                sid = sat[sid]
            except IndexError:
                self.add(
                    oob_code, 'error',
                    '%s: allocation table has no entry for sector %d'
                    % (label, sid),
                    component=component, path=list(path),
                )
                return path
        if sid < 0 and sid != EOCSID:
            self.add(
                terminator_code, 'error',
                '%s chain ends on special sector id %d instead of EOCSID (%d)'
                % (label, sid, EOCSID),
                component=component, path=list(path),
            )
        return path

    # ------------------------------------------------------------------
    # stage 1: header
    # ------------------------------------------------------------------
    def parse_header(self):
        mem = self.mem
        if len(mem) < 512:
            self.add(
                'header.header_truncated', 'error',
                'File is %d bytes long; an OLE2 header needs 512 bytes'
                % len(mem),
            )
            return False
        if mem[0:8] != SIGNATURE:
            self.add(
                'header.bad_signature', 'error',
                'Missing OLE2 signature %r, found %r'
                % (SIGNATURE, bytes(mem[0:8])),
            )
            return False
        if mem[28:30] != b'\xFE\xFF':
            self.add(
                'header.bad_byte_order', 'error',
                'Expected little-endian marker 0xFFFE, found %r'
                % bytes(mem[28:30]),
            )
            return False
        ssz, sssz = unpack('<HH', mem[30:34])
        if ssz > 20:
            self.add(
                'header.bad_sector_size', 'error',
                'Power-of-two sector size exponent %d is implausible' % ssz,
            )
            ssz = 9
        if sssz > ssz:
            self.add(
                'header.bad_sector_size', 'warning',
                'Short sector exponent %d exceeds sector exponent %d; '
                'assuming 64-byte short sectors' % (sssz, ssz),
            )
            sssz = 6
        self.sec_size = 1 << ssz
        self.short_sec_size = 1 << sssz
        (
            _sat_tot_secs, self.dir_first_sec_sid, _unused,
            self.min_size_std_stream, self.ssat_first_sec_sid,
            self.ssat_tot_secs, self.msatx_first_sec_sid,
            self.msatx_tot_secs,
        ) = unpack('<iiiiiiii', mem[44:76])
        data_len = len(mem) - 512
        self.mem_data_secs, leftover = divmod(data_len, self.sec_size)
        if leftover:
            # partial trailing sector still occupies the final slot
            self.mem_data_secs += 1
            self.add(
                'header.file_truncated', 'warning',
                'File size %d is not 512 plus a whole number of %d-byte '
                'sectors' % (len(mem), self.sec_size),
            )
        return True

    # ------------------------------------------------------------------
    # stage 2: MSAT
    # ------------------------------------------------------------------
    def build_msat(self):
        nent = self.sec_size // 4
        fmt = '<%di' % nent
        msat = list(unpack('<109i', self.mem[76:512]))
        if (
            self.msatx_tot_secs == 0 and
            self.msatx_first_sec_sid in (EOCSID, FREESID, 0)
        ):
            return msat
        sid = self.msatx_first_sec_sid
        visited = set()
        while sid not in (EOCSID, FREESID, MSATSID):
            if not self.tick():
                self.add(
                    'msat.budget_exceeded', 'error',
                    'Traversal budget exceeded following the MSAT extension '
                    'chain at sector %d' % sid,
                    component='MSAT', path=sorted(visited),
                )
                break
            if sid >= self.mem_data_secs:
                self.add(
                    'msat.out_of_bounds', 'error',
                    'MSAT extension references sector %d but only %d '
                    'sectors exist' % (sid, self.mem_data_secs),
                    component='MSAT', path=sorted(visited),
                )
                break
            if sid < 0:
                self.add(
                    'msat.bad_sid', 'error',
                    'MSAT extension chain hit invalid sector id %d' % sid,
                    component='MSAT', path=sorted(visited),
                )
                break
            if sid in visited:
                self.add(
                    'msat.cycle', 'error',
                    'MSAT extension chain cycles back to sector %d' % sid,
                    component='MSAT', path=sorted(visited),
                )
                break
            visited.add(sid)
            self._claim(
                self.owners, sid, 'MSAT extension', 'msat.cross_link',
                'MSAT', sorted(visited),
            )
            sector = self._read_sector(sid)
            entries = list(unpack(fmt, sector))
            sid = entries.pop()  # last entry = next MSAT sector id
            msat.extend(entries)
        return msat

    # ------------------------------------------------------------------
    # stage 3: SAT
    # ------------------------------------------------------------------
    def build_sat(self, msat):
        nent = self.sec_size // 4
        fmt = '<%di' % nent
        trunc_warned = False
        sat_sector_path = []
        for msid in msat:
            if msid in (FREESID, EOCSID):
                continue
            if msid >= self.mem_data_secs:
                if not trunc_warned:
                    self.add(
                        'sat.sector_out_of_bounds', 'error',
                        'MSAT points at SAT sector %d but only %d sectors '
                        'exist; the file is truncated or the MSAT is corrupt'
                        % (msid, self.mem_data_secs),
                        component='SAT',
                    )
                    trunc_warned = True
                continue
            if msid < EOCSID:
                self.add(
                    'sat.bad_sid', 'error',
                    'MSAT contains invalid SAT sector id %d' % msid,
                    component='SAT',
                )
                continue
            claimed = self._claim(
                self.owners, msid, 'SAT', 'sat.cross_link', 'SAT',
                list(sat_sector_path),
            )
            if not claimed:
                # Duplicate SAT sector: a cycle in the MSAT-derived list.
                continue
            sat_sector_path.append(msid)
            if not self.tick():
                self.add(
                    'sat.budget_exceeded', 'error',
                    'Traversal budget exceeded while reading SAT sectors',
                    component='SAT',
                )
                break
            sector = self._read_sector(msid)
            self.SAT.extend(unpack(fmt, sector))
        if len(self.SAT) < self.mem_data_secs:
            self.add(
                'sat.truncated', 'error',
                'SAT describes %d sector entries but the file holds %d '
                'sectors' % (len(self.SAT), self.mem_data_secs),
                component='SAT',
            )

    # ------------------------------------------------------------------
    # stage 4: directory
    # ------------------------------------------------------------------
    def parse_directory(self):
        if self.dir_first_sec_sid < 0:
            self.add(
                'directory.missing', 'error',
                'Directory chain start sector id is %d'
                % self.dir_first_sec_sid,
                component='directory',
            )
            return False
        path = self._walk_fat_chain(
            self.dir_first_sec_sid, self.SAT, self.mem_data_secs,
            'directory', self.owners, 'directory.cross_link', 'directory',
            'directory.chain_out_of_bounds', 'directory.chain_cycle',
            'directory.bad_terminator',
        )
        dbytes = b''.join(self._read_sector(sid) for sid in path)
        dirlist = []
        for did, pos in enumerate(range(0, len(dbytes), 128)):
            raw = dbytes[pos:pos + 128]
            if len(raw) < 128:
                self.add(
                    'directory.truncated', 'error',
                    'Directory entry %d is only %d bytes (need 128)'
                    % (did, len(raw)),
                    component='DID %d' % did, path=list(path),
                )
                break
            dirlist.append(self._make_dir_node(did, raw))
        if not dirlist:
            self.add(
                'directory.missing', 'error',
                'No directory entries could be read',
                component='directory',
            )
            return False
        self.dirlist = dirlist
        self._build_tree()
        return True

    def _make_dir_node(self, did, raw):
        (cbufsize, etype, colour, left, right, root) = \
            unpack('<HBBiii', raw[64:80])
        first_sid, tot_size = unpack('<ii', raw[116:124])
        if cbufsize == 0:
            name = ''
        else:
            name = raw[0:max(0, cbufsize - 2)].decode('utf_16_le', 'replace')
        node = _DirNode(did, name, etype, left, right, root,
                        first_sid, tot_size)
        return node

    def _build_tree(self):
        """Iterative replacement for compdoc._build_family_tree recursion."""
        dirlist = self.dirlist
        for node in dirlist:
            node.children = []
            node.parent = -1
        # Bounded explicit stack of (parent_DID, child_DID) work items.
        stack = [(0, dirlist[0].root_DID)]
        # A DID must only be attached once; a self-loop or a DID reached via
        # two parents proves the directory tree is cyclic/corrupt.
        attached = set()
        while stack:
            if not self.tick():
                self.add(
                    'directory.budget_exceeded', 'error',
                    'Traversal budget exceeded while building the directory '
                    'tree',
                    component='directory',
                )
                break
            parent_did, child_did = stack.pop()
            if child_did < 0:
                continue
            if child_did >= len(dirlist):
                self.add(
                    'directory.bad_did', 'error',
                    'Directory reference %d is out of range (only %d '
                    'entries)' % (child_did, len(dirlist)),
                    component='DID %d' % parent_did,
                )
                continue
            if child_did in attached:
                self.add(
                    'directory.tree_cycle', 'error',
                    'Directory entry %d is attached more than once '
                    '(first under parent %d, again under parent %d)'
                    % (child_did,
                       dirlist[child_did].parent, parent_did),
                    component='DID %d' % child_did,
                )
                continue
            attached.add(child_did)
            dirlist[parent_did].children.append(child_did)
            dirlist[child_did].parent = parent_did
            child = dirlist[child_did]
            stack.append((parent_did, child.left_DID))
            stack.append((parent_did, child.right_DID))
            if child.etype == 1:  # storage: descend into its red-black tree
                stack.append((child_did, child.root_DID))
        self._assign_paths()

    def _assign_paths(self):
        # BFS from the root storage so every name gets a readable full path.
        dirlist = self.dirlist
        names = {0: ''}
        queue = [0]
        head = 0
        while head < len(queue):
            if not self.tick():
                break
            did = queue[head]
            head += 1
            prefix = names.get(did, '')
            for child in dirlist[did].children:
                child_node = dirlist[child]
                full = prefix + '/' + child_node.name if prefix \
                    else '/' + child_node.name
                names[child] = full
                queue.append(child)
        for node in dirlist:
            node.path = names.get(node.DID, node.name)
            self.result.directory_entries.append((
                node.DID, node.name, node.etype, node.first_SID,
                node.tot_size, node.path,
            ))
        root = dirlist[0]
        if root.etype != 5:
            self.add(
                'directory.bad_root', 'error',
                'DID 0 has entry type %d, expected 5 (root storage)'
                % root.etype,
                component=root.path or 'Root Entry',
            )

    # ------------------------------------------------------------------
    # stage 5: SSCS (mini stream container) and SSAT
    # ------------------------------------------------------------------
    def build_mini_fat(self):
        root = self.dirlist[0]
        if root.first_SID < 0 or root.tot_size == 0:
            self.SSCS = b''
            if self.ssat_tot_secs > 0:
                self.add(
                    'ssat.empty_with_ssat', 'warning',
                    'SSAT declares %d sector(s) but the root mini stream '
                    'is empty' % self.ssat_tot_secs,
                    component=root.path or 'Root Entry',
                )
            return
        sscs_path = self._walk_fat_chain(
            root.first_SID, self.SAT, self.mem_data_secs,
            'SSCS (root mini stream)', self.owners, 'sscs.cross_link',
            root.path or 'Root Entry',
            'sscs.chain_out_of_bounds', 'sscs.chain_cycle',
            'sscs.bad_terminator',
        )
        actual = len(sscs_path) * self.sec_size
        if actual < root.tot_size:
            self.add(
                'sscs.size_mismatch', 'error',
                'Root mini stream declares %d bytes but its chain only '
                'holds %d bytes' % (root.tot_size, actual),
                component=root.path or 'Root Entry', path=list(sscs_path),
            )
        self.SSCS = b''.join(self._read_sector(sid) for sid in sscs_path)

        nent = self.sec_size // 4
        fmt = '<%di' % nent
        ssat_path = self._walk_fat_chain(
            self.ssat_first_sec_sid, self.SAT, self.mem_data_secs,
            'SSAT', self.owners, 'ssat.cross_link', 'SSAT',
            'ssat.chain_out_of_bounds', 'ssat.chain_cycle',
            'ssat.bad_terminator',
        )
        if len(ssat_path) != self.ssat_tot_secs:
            self.add(
                'ssat.size_mismatch', 'error',
                'SSAT declares %d sector(s) but its FAT chain provides %d'
                % (self.ssat_tot_secs, len(ssat_path)),
                component='SSAT',
                path=list(ssat_path),
            )
        for sid in ssat_path:
            sector = self._read_sector(sid)
            self.SSAT.extend(unpack(fmt, sector))
        self.mini_sectors = max(
            0, (len(self.SSCS) + self.short_sec_size - 1)
            // self.short_sec_size
        )

    # ------------------------------------------------------------------
    # stage 6: every user stream
    # ------------------------------------------------------------------
    def analyse_streams(self):
        for node in self.dirlist:
            if not self.tick():
                self.add(
                    'stream.budget_exceeded', 'error',
                    'Traversal budget exceeded before all streams were '
                    'checked (stopped at %r)' % node.path,
                    component=node.path,
                )
                break
            if node.DID == 0:
                continue  # root entry handled in build_mini_fat
            if node.etype == 0:
                continue  # unused directory slot
            if node.etype == 1:
                if node.first_SID >= 0 or node.tot_size:
                    self.add(
                        'stream.storage_has_chain', 'warning',
                        'Storage entry declares first_SID=%d size=%d; both '
                        'should be zero for a storage'
                        % (node.first_SID, node.tot_size),
                        component=node.path,
                    )
                continue
            if node.etype != 2:
                self.add(
                    'directory.unknown_etype', 'warning',
                    'Entry %r has unknown object type %d'
                    % (node.path, node.etype),
                    component=node.path,
                )
                continue
            self._analyse_user_stream(node)

    def _analyse_user_stream(self, node):
        label = 'stream %r' % node.path
        if node.tot_size == 0:
            if node.first_SID not in (EOCSID, FREESID):
                self.add(
                    'stream.bad_first_sid', 'warning',
                    'Zero-size stream %r should start on EOCSID/FREESID, '
                    'not sector %d' % (node.path, node.first_SID),
                    component=node.path,
                )
            return
        if node.first_SID < 0:
            self.add(
                'stream.bad_first_sid', 'error',
                'Stream %r declares %d bytes but its first sector id is %d'
                % (node.path, node.tot_size, node.first_SID),
                component=node.path,
            )
            return
        if node.tot_size >= self.min_size_std_stream:
            self._analyse_standard_stream(node, label)
        else:
            self._analyse_mini_stream(node, label)

    def _analyse_standard_stream(self, node, label):
        path = self._walk_fat_chain(
            node.first_SID, self.SAT, self.mem_data_secs,
            label, self.owners, 'stream.cross_link', node.path,
            'stream.chain_out_of_bounds', 'stream.chain_cycle',
            'stream.bad_terminator',
        )
        node.chain_path = path
        node.chain_kind = 'standard'
        expected_sectors = (
            (node.tot_size + self.sec_size - 1) // self.sec_size
        )
        if len(path) != expected_sectors:
            self.add(
                'stream.size_mismatch', 'error',
                'Stream %r declares %d bytes (%d sector(s)) but its FAT '
                'chain provides %d sector(s)'
                % (node.path, node.tot_size, expected_sectors, len(path)),
                component=node.path, path=list(path),
            )

    def _analyse_mini_stream(self, node, label):
        if not self.SSCS:
            self.add(
                'mini.no_container', 'error',
                'Mini stream %r exists but the root mini stream (SSCS) is '
                'empty' % node.path,
                component=node.path,
            )
            return
        path = self._walk_fat_chain(
            node.first_SID, self.SSAT, self.mini_sectors,
            label + ' (mini)', self.mini_owners, 'mini.cross_link',
            node.path,
            'mini.chain_out_of_bounds', 'mini.chain_cycle',
            'mini.bad_terminator',
        )
        node.chain_path = path
        node.chain_kind = 'mini'
        expected_sectors = (
            (node.tot_size + self.short_sec_size - 1) // self.short_sec_size
        )
        if len(path) != expected_sectors:
            self.add(
                'mini.size_mismatch', 'error',
                'Mini stream %r declares %d bytes (%d short sector(s)) but '
                'its miniFAT chain provides %d sector(s)'
                % (node.path, node.tot_size, expected_sectors, len(path)),
                component=node.path, path=list(path),
            )


class _DirNode(object):
    __slots__ = (
        'DID', 'name', 'etype', 'left_DID', 'right_DID', 'root_DID',
        'first_SID', 'tot_size', 'children', 'parent', 'path',
        'chain_path', 'chain_kind',
    )

    def __init__(self, did, name, etype, left, right, root,
                 first_sid, tot_size):
        self.DID = did
        self.name = name
        self.etype = etype
        self.left_DID = left
        self.right_DID = right
        self.root_DID = root
        self.first_SID = first_sid
        self.tot_size = tot_size
        self.children = []
        self.parent = -1
        self.path = name
        self.chain_path = []
        self.chain_kind = None


#: BOF opcodes recognised by the BIFF parser (see xlrd.biffh.bofcodes).
_BOF_OPCODES = (0x0809, 0x0409, 0x0209, 0x0009)
_BOF_LENGTHS = (4, 6, 8, 16, 20)


def _check_biff(stream):
    """
    Validate only the leading BOF record of a located workbook stream.

    Returns ``(state, detail)`` where state is ``"biff_ok"`` or
    ``"biff_invalid"``.  This mirrors the first decision Book.getbof makes
    without attempting a full (repair-tolerant) parse.
    """
    if len(stream) < 4:
        return 'biff_invalid', (
            'Workbook stream is only %d byte(s); a BOF record needs at '
            'least 4 bytes' % len(stream)
        )
    opcode = stream[0] | (stream[1] << 8)
    length = stream[2] | (stream[3] << 8)
    if opcode not in _BOF_OPCODES:
        return 'biff_invalid', (
            'First BIFF record has opcode 0x%04X, expected one of %s'
            % (opcode, ', '.join('0x%04X' % c for c in _BOF_OPCODES))
        )
    if not (4 <= length <= 20):
        return 'biff_invalid', (
            'BOF record length %d is outside the supported range 4..20'
            % length
        )
    if length not in _BOF_LENGTHS:
        return 'biff_invalid', (
            'BOF record length %d is not a known BIFF BOF size %r'
            % (length, list(_BOF_LENGTHS))
        )
    if len(stream) < 4 + length:
        return 'biff_invalid', (
            'BOF record declares %d payload bytes but only %d remain'
            % (length, len(stream) - 4)
        )
    return 'biff_ok', 'Leading BOF record is well-formed'


def _analyse(mem, budget):
    analyzer = _Analyzer(mem, budget)
    result = analyzer.result

    if not analyzer.parse_header():
        result.can_locate_workbook = False
        result.reason = (
            'OLE header is invalid; the container cannot be navigated, so '
            'the BIFF stream was never reached'
        )
        return result

    msat = analyzer.build_msat()
    analyzer.build_sat(msat)

    if not analyzer.parse_directory():
        result.can_locate_workbook = False
        result.reason = (
            'Directory chain is unreadable; no Workbook entry can be '
            'located, so the BIFF stream was never reached'
        )
        return result

    analyzer.build_mini_fat()
    analyzer.analyse_streams()

    workbook = _find_workbook(analyzer)
    if workbook is None:
        result.can_locate_workbook = False
        result.reason = (
            'No Workbook or Book directory entry exists; nothing to hand '
            'to the BIFF parser'
        )
        return result

    result.workbook_name = workbook.path
    result.workbook_stream_type = workbook.chain_kind
    result.workbook_path = list(workbook.chain_path)

    blocking = [
        issue for issue in result.issues
        if issue.severity == 'error' and _issue_blocks_workbook(issue, workbook)
    ]
    if blocking:
        result.can_locate_workbook = False
        codes = sorted(set(issue.code for issue in blocking))
        result.reason = (
            'Container is damaged before/within the Workbook chain (%s); '
            'the BIFF stream cannot be proven intact and was not parsed'
            % ', '.join(codes)
        )
        return result

    result.can_locate_workbook = True
    if workbook.chain_kind is None:
        result.can_locate_workbook = False
        result.workbook_stream_type = None
        result.reason = (
            'Workbook entry %r declares %d bytes but has no sector chain to '
            'read' % (workbook.path, workbook.tot_size)
        )
        return result
    stream = _extract_workbook(analyzer, workbook)
    if stream is None:
        result.can_locate_workbook = False
        result.reason = (
            'Workbook chain points outside the file; its bytes cannot be '
            'read'
        )
        return result
    state, detail = _check_biff(stream)
    result.biff_state = state
    result.biff_detail = detail
    if state == 'biff_ok':
        result.reason = (
            'OLE container is structurally sound and the Workbook stream '
            'starts with a valid BIFF BOF record; open_workbook may proceed '
            'to BIFF parsing'
        )
    else:
        result.reason = (
            'OLE container is structurally sound but the located Workbook '
            'stream is not valid BIFF: %s' % detail
        )
    return result


def _find_workbook(analyzer):
    wanted = ('workbook', 'book')
    for node in analyzer.dirlist:
        if node.etype == 2 and node.name.lower() in wanted:
            return node
    return None


def _issue_blocks_workbook(issue, workbook):
    """
    Decide whether an issue touches the structures required to locate and
    read the specific workbook stream.

    Damage confined to an unrelated user stream is still reported, but it
    does not by itself make the Workbook unreadable.
    """
    if issue.code in _FATAL_CODES:
        # Header/MSAT/SAT/directory/SSAT/SSCS damage undermines every chain.
        return True
    if issue.component == workbook.path:
        return True
    if issue.component in ('Workbook', 'Book'):
        return True
    return False


def _extract_workbook(analyzer, node):
    try:
        if node.chain_kind == 'standard':
            parts = [
                analyzer._read_sector(sid) for sid in node.chain_path
            ]
            return b''.join(parts)[:node.tot_size]
        if node.chain_kind == 'mini':
            parts = []
            for sid in node.chain_path:
                start = sid * analyzer.short_sec_size
                parts.append(
                    analyzer.SSCS[start:start + analyzer.short_sec_size]
                )
            return b''.join(parts)[:node.tot_size]
    except (IndexError, MemoryError):
        return None
    return None


def diagnose_workbook_container(filename=None, file_contents=None,
                                traversal_budget=DEFAULT_TRAVERSAL_BUDGET):
    """
    Read-only OLE2 container diagnosis for an XLS file.

    The container is checked *before* any BIFF parsing.  Nothing is repaired,
    rewritten or skipped: every structural defect found is returned in
    :attr:`ContainerDiagnosis.issues`, and :func:`open_workbook` keeps its
    original exception contract.

    :param filename:
      Path of the ``.xls`` file to diagnose.  ``~`` is expanded.  Ignored
      when ``file_contents`` is supplied.
    :param file_contents:
      Raw file bytes (or any object supporting slicing and ``len``).
    :param traversal_budget:
      Maximum number of chain-walk steps across the whole diagnosis, so a
      cyclic or otherwise malicious directory can never loop forever.
    :returns:
      A :class:`ContainerDiagnosis`.  Check
      :attr:`ContainerDiagnosis.can_locate_workbook` and
      :attr:`ContainerDiagnosis.reason` for the go/no-go decision and
      :attr:`ContainerDiagnosis.biff_state` to distinguish a sound container
      with invalid BIFF from pre-BIFF container damage.
    """
    if traversal_budget <= 0:
        raise ValueError(
            'traversal_budget must be a positive integer, got %r'
            % traversal_budget
        )
    if file_contents is None:
        if filename is None:
            raise TypeError(
                'diagnose_workbook_container needs either filename or '
                'file_contents'
            )
        with open(os.path.expanduser(filename), 'rb') as handle:
            file_contents = handle.read()
    try:
        return _analyse(file_contents, traversal_budget)
    except Exception as exc:  # defensive: diagnostics must never crash callers
        result = ContainerDiagnosis()
        result.can_locate_workbook = False
        result.reason = (
            'Diagnosis aborted on an unexpected %s: %s; the container '
            'cannot be trusted' % (type(exc).__name__, exc)
        )
        result.issues.append(
            Issue(
                'diagnoser.aborted', 'error', result.reason,
            )
        )
        return result
