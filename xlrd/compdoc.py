# Copyright (c) 2005-2012 Stephen John Machin, Lingfo Pty Ltd
# This module is part of the xlrd package, which is released under a
# BSD-style licence.
# No part of the content of this file was derived from the works of
# David Giffin.
"""
Implements the minimal functionality required
to extract a "Workbook" or "Book" stream (as one big string)
from an OLE2 Compound Document file.
"""
import array
import mmap
from bisect import bisect_left
from struct import unpack
from logging import getLogger

from .timemachine import *


logger = getLogger(__name__)
#: Magic cookie that should appear in the first 8 bytes of the file.
SIGNATURE = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"

EOCSID = -2
FREESID = -1
SATSID = -3
MSATSID = -4
EVILSID = -5


class CompDocError(Exception):
    pass


class DirNode(object):
    def __init__(self, DID, dent):
        # dent is the 128-byte directory entry
        self.DID = DID
        (cbufsize, self.etype, self.colour, self.left_DID, self.right_DID,
         self.root_DID) = \
            unpack('<HBBiii', dent[64:80])
        (self.first_SID, self.tot_size) = \
            unpack('<ii', dent[116:124])
        if cbufsize == 0:
            self.name = UNICODE_LITERAL('')
        else:
            self.name = unicode(dent[0:cbufsize - 2], 'utf_16_le')  # omit the trailing U+0000
        self.children = []  # filled in later
        self.parent = -1  # indicates orphan; fixed up later
        self.tsinfo = unpack('<IIII', dent[100:116])

    def __str__(self):
        msg = f"DID={self.DID} name={self.name} etype={self.etype} DIDs(left={self.left_DID} right={self.right_DID} " \
              f"root='{self.root_DID}' parent={self.parent} kids={self.children}) first_SID={self.first_SID} " \
              f"tot_size={self.tot_size}\n" \
              f"timestamp info: {self.tsinfo}"
        return msg


def _build_family_tree(dirlist, parent_DID, child_DID):
    if child_DID < 0:
        return
    _build_family_tree(dirlist, parent_DID, dirlist[child_DID].left_DID)
    dirlist[parent_DID].children.append(child_DID)
    dirlist[child_DID].parent = parent_DID
    _build_family_tree(dirlist, parent_DID, dirlist[child_DID].right_DID)
    if dirlist[child_DID].etype == 1:  # storage
        _build_family_tree(dirlist, child_DID, dirlist[child_DID].root_DID)


class ScatteredMemory:
    """Return a slice of big file with mmap"""

    def __init__(self, filename, slices):
        self.filename = filename
        self.mem = b''
        self.slices = slices

        sizes = [y - x for x, y in slices]
        self.cum_sum = [sum(sizes[:x + 1]) for x in range(0, len(sizes))]

    @staticmethod
    def _mmap_closed(mmap):
        return mmap.closed

    def __getitem__(self, item):
        if not self.mem or self._mmap_closed(self.mem):
            with open(self.filename) as fi:
                self.mem = mmap.mmap(fi.fileno(), 0, access=mmap.ACCESS_READ)

        if type(item) is int:
            idx = bisect_left(self.cum_sum, item)
            _item = item - (self.cum_sum[idx - 1] if idx > 0 else 0)
            _idx = self.slices[idx][0] + _item
            return self.mem[_idx]

        else:
            start, end = item.start, item.stop
            length = end - start
            assert not item.step, "stepping is not supported"
            idx_0 = bisect_left(self.cum_sum, start)
            idx_1 = bisect_left(self.cum_sum, end)

            d_start = 0 if not idx_0 else self.cum_sum[idx_0 - 1]

            data = b''.join(self.mem[start_pos:end_pos] for start_pos, end_pos in self.slices[idx_0:idx_1 + 1])
            _start = start - d_start
            mem = data[_start:_start + length]
            del data
            return mem


class CompDoc(object):
    """
    Compound document handler.

    :param mem:
      The raw contents of the file, as a string, or as an :class:`mmap.mmap`
      object. The only operation it needs to support is slicing.
    """

    def __init__(self, xls_path, mem, ignore_workbook_corruption=False):
        self.xls_path = xls_path
        self.ignore_workbook_corruption = ignore_workbook_corruption

        if mem[0:8] != SIGNATURE:
            raise CompDocError('Not an OLE2 compound document')
        if mem[28:30] != b'\xFE\xFF':
            raise CompDocError('Expected "little-endian" marker, found %r' % mem[28:30])
        revision, version = unpack('<HH', mem[24:28])
        logger.debug(f"CompDoc format: version=0x{version:04x} revision=0x{revision:04x}")

        self.mem = mem
        ssz, sssz = unpack('<HH', mem[30:34])
        if ssz > 20:  # allows for 2**20 bytes i.e. 1MB
            logger.warning(f"sector size (2**{ssz}) is preposterous; assuming 512 and continuing ...")
            ssz = 9
        if sssz > ssz:
            logger.warning(f"short stream sector size (2**{sssz}) is preposterous; assuming 64 and continuing ...")
            sssz = 6
        self.sec_size = sec_size = 1 << ssz
        self.short_sec_size = 1 << sssz

        if self.sec_size != 512 or self.short_sec_size != 64:
            logger.info(f"@@@@ sec_size={self.sec_size} short_sec_size={self.short_sec_size}")
        (
            SAT_tot_secs, self.dir_first_sec_sid, _unused, self.min_size_std_stream,
            SSAT_first_sec_sid, SSAT_tot_secs,
            MSATX_first_sec_sid, MSATX_tot_secs,
        ) = unpack('<iiiiiiii', mem[44:76])
        mem_data_len = len(mem) - 512
        mem_data_secs, left_over = divmod(mem_data_len, sec_size)

        if left_over:
            # raise CompDocError("Not a whole number of sectors")
            mem_data_secs += 1
            logger.warning(f"file size ({len(mem)}) not 512 + multiple of sector size ({sec_size})")

        self.mem_data_secs = mem_data_secs  # use for checking later
        self.mem_data_len = mem_data_len
        seen = self.seen = array.array('B', [0]) * mem_data_secs

        # logger.debug(f'sec sizes', ssz, sssz, sec_size, self.short_sec_size)
        # logger.debug(f"mem data: {mem_data_len} bytes == {mem_data_secs} sectors")
        # logger.debug(f"SAT_tot_secs={SAT_tot_secs}, dir_first_sec_sid={self.dir_first_sec_sid}, "
        #              f"min_size_std_stream={self.min_size_std_stream}")
        # logger.debug(f"SSAT_first_sec_sid={SSAT_first_sec_sid}, SSAT_tot_secs={SSAT_tot_secs}")
        # logger.debug(f"MSATX_first_sec_sid={MSATX_first_sec_sid}, MSATX_tot_secs={MSATX_tot_secs}")

        nent = sec_size // 4  # number of SID entries in a sector
        fmt = "<%di" % nent
        trunc_warned = 0
        #
        # === build the MSAT ===
        #
        MSAT = list(unpack('<109i', mem[76:512]))
        SAT_sectors_reqd = (mem_data_secs + nent - 1) // nent
        expected_MSATX_sectors = max(0, (SAT_sectors_reqd - 109 + nent - 2) // (nent - 1))
        actual_MSATX_sectors = 0
        if MSATX_tot_secs == 0 and MSATX_first_sec_sid in (EOCSID, FREESID, 0):
            # Strictly, if there is no MSAT extension, then MSATX_first_sec_sid
            # should be set to EOCSID ... FREESID and 0 have been met in the wild.
            pass  # Presuming no extension
        else:
            sid = MSATX_first_sec_sid
            while sid not in (EOCSID, FREESID, MSATSID):
                # Above should be only EOCSID according to MS & OOo docs
                # but Excel doesn't complain about FREESID. Zero is a valid
                # sector number, not a sentinel.
                logger.debug(f'MSATX: sid={sid} (0x{sid:08X})')
                if sid >= mem_data_secs:
                    msg = "MSAT extension: accessing sector %d but only %d in file" % (sid, mem_data_secs)
                    raise CompDocError(msg)
                elif sid < 0:
                    raise CompDocError(f"MSAT extension: invalid sector id: {sid}")
                if seen[sid]:
                    raise CompDocError(f"MSAT corruption: seen[{sid}] == {seen[sid]}")

                seen[sid] = 1
                actual_MSATX_sectors += 1

                if actual_MSATX_sectors > expected_MSATX_sectors:
                    logger.debug(f"[1]===>>> {mem_data_secs}, {nent}, {SAT_sectors_reqd}, "
                                 f"{expected_MSATX_sectors}, {actual_MSATX_sectors}")
                offset = 512 + sec_size * sid
                MSAT.extend(unpack(fmt, mem[offset:offset + sec_size]))
                sid = MSAT.pop()  # last sector id is sid of next sector in the chain

        if actual_MSATX_sectors != expected_MSATX_sectors:
            logger.debug(f"[2]===>>> {mem_data_secs}, {nent}, {SAT_sectors_reqd}, "
                         f"{expected_MSATX_sectors}, {actual_MSATX_sectors}")

        # dump_list(MSAT, 10, header=f"MSAT: len={len(MSAT)}")
        #
        # === build the SAT ===
        #
        self.SAT = []
        actual_SAT_sectors = 0
        dump_again = 0
        for msidx, msid in enumerate(MSAT):
            if msid in (FREESID, EOCSID):
                # Specification: the MSAT array may be padded with trailing FREESID entries.
                # Toleration: a FREESID or EOCSID entry anywhere in the MSAT array will be ignored.
                continue
            if msid >= mem_data_secs:
                if not trunc_warned:
                    logger.warning(f"File is truncated, or OLE2 MSAT is corrupt!!")
                    logger.warning(f"Trying to access sector {msid} but only {mem_data_secs} available")
                    trunc_warned = 1
                MSAT[msidx] = EVILSID
                dump_again = 1
                continue
            elif msid < -2:
                raise CompDocError(f"MSAT: invalid sector id: {msid}")
            if seen[msid]:
                raise CompDocError(f"MSAT extension corruption: seen[{msid}] == {seen[msid]}")
            seen[msid] = 2
            actual_SAT_sectors += 1
            if actual_SAT_sectors > SAT_sectors_reqd:
                logger.debug(f"[3]===>>> {mem_data_secs}, {nent}, {SAT_sectors_reqd}, {expected_MSATX_sectors}, "
                             f"{actual_MSATX_sectors}, {actual_SAT_sectors}, {msid}")
            offset = 512 + sec_size * msid
            self.SAT.extend(unpack(fmt, mem[offset:offset + sec_size]))

            # dump_list(self.SAT, 10, header=f"SAT: len={len(self.SAT)}")

        if dump_again:
            # dump_list(MSAT, 10, header=f"MSAT: len={len(MSAT)}")
            for satx in xrange(mem_data_secs, len(self.SAT)):
                self.SAT[satx] = EVILSID
            # dump_list(self.SAT, 10, header=f"SAT: len={len(self.SAT)}")
        #
        # === build the directory ===
        #
        dbytes = self._get_stream(self.mem, 512, self.SAT, self.sec_size,
                                  self.dir_first_sec_sid, name="directory", seen_id=3)
        dirlist = []
        did = -1
        for pos in xrange(0, len(dbytes), 128):
            did += 1
            dirlist.append(DirNode(did, dbytes[pos:pos + 128]))
        self.dirlist = dirlist
        _build_family_tree(dirlist, 0, dirlist[0].root_DID)  # and stand well back ...
        # for d in dirlist:
        #     logger.debug(str(d))
        #
        # === get the SSCS ===
        #
        sscs_dir = self.dirlist[0]
        assert sscs_dir.etype == 5  # root entry
        if sscs_dir.first_SID < 0 or sscs_dir.tot_size == 0:
            # Problem reported by Frank Hoffsuemmer: some software was
            # writing -1 instead of -2 (EOCSID) for the first_SID
            # when the SCCS was empty. Not having EOCSID caused assertion
            # failure in _get_stream.
            # Solution: avoid calling _get_stream in any case when the
            # SCSS appears to be empty.
            self.SSCS = ""
        else:
            self.SSCS = self._get_stream(self.mem, 512, self.SAT, sec_size,
                                         sscs_dir.first_SID, sscs_dir.tot_size, name="SSCS", seen_id=4)
        # if DEBUG: print >> logfile, "SSCS", repr(self.SSCS)
        #
        # === build the SSAT ===
        #
        self.SSAT = []
        if SSAT_tot_secs > 0 and sscs_dir.tot_size == 0:
            logger.warning(f"OLE2 inconsistency: SSCS size is 0 but SSAT size is non-zero")

        if sscs_dir.tot_size > 0:
            sid = SSAT_first_sec_sid
            nsecs = SSAT_tot_secs
            while sid >= 0 and nsecs > 0:
                if seen[sid]:
                    raise CompDocError(f"SSAT corruption: seen[{sid}] == {seen[sid]}")
                seen[sid] = 5
                nsecs -= 1
                start_pos = 512 + sid * sec_size
                news = list(unpack(fmt, mem[start_pos:start_pos + sec_size]))
                self.SSAT.extend(news)
                sid = self.SAT[sid]

            logger.debug(f"SSAT last sid {sid}; remaining sectors {nsecs}")
            assert nsecs == 0 and sid == EOCSID

        # dump_list(self.SSAT, 10, header="SSAT")
        # dump_list(seen, 20, header="seen")

    def _get_stream(self, mem, base, sat, sec_size, start_sid, size=None, name='', seen_id=None):
        # print >> self.logfile, "_get_stream", base, sec_size, start_sid, size
        sectors = []
        s = start_sid
        if size is None:
            # nothing to check against
            while s >= 0:
                if seen_id is not None:
                    if self.seen[s]:
                        raise CompDocError(f"{name} corruption: seen[{s}] == {self.seen[s]}")
                    self.seen[s] = seen_id
                start_pos = base + s * sec_size
                sectors.append(mem[start_pos:start_pos + sec_size])
                try:
                    s = sat[s]
                except IndexError:
                    raise CompDocError(f"OLE2 stream {name}: sector allocation table invalid entry ({s})")
            assert s == EOCSID
        else:
            todo = size
            while s >= 0:
                if seen_id is not None:
                    if self.seen[s]:
                        raise CompDocError(f"{name} corruption: seen[{s}] == {self.seen[s]}")
                    self.seen[s] = seen_id
                start_pos = base + s * sec_size
                grab = sec_size
                if grab > todo:
                    grab = todo
                todo -= grab
                sectors.append(mem[start_pos:start_pos + grab])
                try:
                    s = sat[s]
                except IndexError:
                    raise CompDocError(f"OLE2 stream {name}: sector allocation table invalid entry ({s})")
            assert s == EOCSID
            if todo != 0:
                logger.warning(f"OLE2 stream {name}: expected size {size}, actual size {size - todo}")

        return b''.join(sectors)

    def _dir_search(self, path, storage_DID=0):
        # Return matching DirNode instance, or None
        head = path[0]
        tail = path[1:]
        dl = self.dirlist
        for child in dl[storage_DID].children:
            if dl[child].name.lower() == head.lower():
                et = dl[child].etype
                if et == 2:
                    return dl[child]
                if et == 1:
                    if not tail:
                        raise CompDocError("Requested component is a 'storage'")
                    return self._dir_search(tail, child)
                logger.debug(str(dl[child]))
                raise CompDocError("Requested stream is not a 'user stream'")
        return None

    def get_named_stream(self, qname):
        """
        Interrogate the compound document's directory; return the stream as a
        string if found, otherwise return ``None``.

        :param qname:
          Name of the desired stream e.g. ``'Workbook'``.
          Should be in Unicode or convertible thereto.
        """
        d = self._dir_search(qname.split("/"))
        if d is None:
            return None
        if d.tot_size >= self.min_size_std_stream:
            return self._get_stream(
                self.mem, 512, self.SAT, self.sec_size, d.first_SID,
                d.tot_size, name=qname, seen_id=d.DID + 6)
        else:
            return self._get_stream(
                self.SSCS, 0, self.SSAT, self.short_sec_size, d.first_SID,
                d.tot_size, name=qname + " (from SSCS)", seen_id=None)

    def locate_named_stream(self, qname):
        """
        Interrogate the compound document's directory.

        If the named stream is not found, ``(None, 0, 0)`` will be returned.

        If the named stream is found and is contiguous within the original
        byte sequence (``mem``) used when the document was opened,
        then ``(mem, offset_to_start_of_stream, length_of_stream)`` is returned.

        Otherwise a new string is built from the fragments and
        ``(new_string, 0, length_of_stream)`` is returned.

        :param qname:
          Name of the desired stream e.g. ``'Workbook'``.
          Should be in Unicode or convertible thereto.
        """
        d = self._dir_search(qname.split("/"))
        if d is None:
            return None, 0, 0
        if d.tot_size > self.mem_data_len:
            raise CompDocError(f"{qname} stream length ({d.tot_size}bytes) > file data size ({self.mem_data_len}bytes)")
        if d.tot_size >= self.min_size_std_stream:
            result = self._locate_stream(
                self.mem, 512, self.SAT, self.sec_size, d.first_SID,
                d.tot_size, qname, d.DID + 6)
            # dump_list(self.seen, 20, header='seen'.center(30))
            return result
        else:
            stream = self._get_stream(self.SSCS, 0, self.SSAT, self.short_sec_size,
                                      d.first_SID, d.tot_size, qname + " (from SSCS)", None)
            return stream, 0, d.tot_size

    def _locate_stream(self, mem, base, sat, sec_size, start_sid, expected_stream_size, qname, seen_id):
        # print >> self.logfile, "_locate_stream", base, sec_size, start_sid, expected_stream_size
        s = start_sid
        if s < 0:
            raise CompDocError(f"_locate_stream: start_sid ({start_sid}) is -ve")
        p = -99  # dummy previous SID
        start_pos = -9999
        end_pos = -8888
        slices = []
        tot_found = 0
        found_limit = (expected_stream_size + sec_size - 1) // sec_size
        while s >= 0:
            if self.seen[s]:
                if not self.ignore_workbook_corruption:
                    dump_list(self.seen, 20, header=f"_locate_stream({qname}): seen")
                    raise CompDocError(f"{qname} corruption: seen[{s}] == {self.seen[s]}")
            self.seen[s] = seen_id
            tot_found += 1
            if tot_found > found_limit:
                # Note: expected size rounded up to higher sector
                raise CompDocError(f"{qname}: size exceeds expected {found_limit * sec_size} bytes; corrupt?")
            if s == p + 1:
                # contiguous sectors
                end_pos += sec_size
            else:
                # start new slice
                if p >= 0:
                    # not first time
                    slices.append((start_pos, end_pos))
                start_pos = base + s * sec_size
                end_pos = start_pos + sec_size
            p = s
            s = sat[s]
        assert s == EOCSID
        assert tot_found == found_limit
        # print >> self.logfile, "_locate_stream(%s): seen" % qname; dump_list(self.seen, 20, self.logfile)
        if not slices:
            # The stream is contiguous ... just what we like!
            return mem, start_pos, expected_stream_size

        slices.append((start_pos, end_pos))
        # print >> self.logfile, "+++>>> %d fragments" % len(slices)
        # return b''.join(mem[start_pos:end_pos] for start_pos, end_pos in slices), 0, expected_stream_size
        if expected_stream_size < 1024 * 1024 * 90:  # less than 90MB, faster to read as a whole than mmaping
            new_mem = b''.join(mem[start_pos:end_pos] for start_pos, end_pos in slices)
        else:
            new_mem = ScatteredMemory(self.xls_path, slices)
        return new_mem, 0, expected_stream_size


# ==========================================================================================
def dump_list(alist, stride:int, header:str=''):
    def _dump_line(d_pos, equal=False):
        msg = [f"{d_pos:5d} " + '=' * equal]
        for value in alist[d_pos:d_pos + stride]:
            msg.append(str(value))
        logger.debug(' '.join(msg))

    if header:
        logger.debug(header)

    pos = None
    old_pos = None
    for pos in xrange(0, len(alist), stride):
        if not old_pos:
            _dump_line(pos)
            old_pos = pos

        elif alist[pos:pos + stride] != alist[old_pos:old_pos + stride]:
            if pos - old_pos > stride:
                _dump_line(pos - stride, equal=True)
            _dump_line(pos)
            old_pos = pos

    if old_pos and pos and pos != old_pos:
        _dump_line(pos, equal=True)
