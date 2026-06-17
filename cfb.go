// cfb.go implements a reader for the Compound File Binary format ([MS-CFB]),
// the OLE2 structured storage container used by .doc, .xls and .ppt files.
//
// Referenced sections of [MS-CFB]:
//   - 2.2 Compound File Header
//   - 2.3 Compound File Sector Numbers and Types (FAT/DIFAT)
//   - 2.4 Compound File FAT Sectors
//   - 2.6 Compound File Directory Sectors
//   - 2.8 Compound File Mini FAT Sectors

package oletext

import (
	"encoding/binary"
	"errors"
	"fmt"
	"unicode/utf16"
)

// Special sector numbers ([MS-CFB] 2.3).
const (
	secMaxRegular = 0xFFFFFFFA // MAXREGSECT
	secDIFAT      = 0xFFFFFFFC // DIFSECT
	secFAT        = 0xFFFFFFFD // FATSECT
	secEndOfChain = 0xFFFFFFFE // ENDOFCHAIN
	secFree       = 0xFFFFFFFF // FREESECT
)

var cfbSignature = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

// dirEntry is a parsed directory entry ([MS-CFB] 2.6.1, 128 bytes each).
type dirEntry struct {
	name      string
	objType   byte // 0=unknown 1=storage 2=stream 5=root
	startSect uint32
	size      uint64
}

// cfbFile is a parsed compound file ready for stream lookups.
type cfbFile struct {
	data       []byte
	sectorSize int
	miniCutoff uint32
	fat        []uint32
	miniFAT    []uint32
	dirs       []dirEntry
	miniStream []byte // contents of the root entry's stream, chained by the mini FAT
}

// parseCFB parses the compound file header, FAT, directory and mini FAT
// from a complete file image.
func parseCFB(data []byte) (*cfbFile, error) {
	if len(data) < 512 {
		return nil, errors.New("file too small to be a compound file")
	}
	for i, b := range cfbSignature {
		if data[i] != b {
			return nil, errors.New("not an OLE2 compound file (bad signature)")
		}
	}
	major := binary.LittleEndian.Uint16(data[26:])
	sectorShift := binary.LittleEndian.Uint16(data[30:])
	// [MS-CFB] 2.2: version 3 -> 512-byte sectors, version 4 -> 4096-byte sectors.
	if !(major == 3 && sectorShift == 9) && !(major == 4 && sectorShift == 12) {
		return nil, fmt.Errorf("unsupported CFB version %d / sector shift %d", major, sectorShift)
	}
	f := &cfbFile{
		data:       data,
		sectorSize: 1 << sectorShift,
		miniCutoff: binary.LittleEndian.Uint32(data[56:]),
	}

	// A well-formed file has at most one FAT/DIFAT sector per stored sector,
	// so the on-disk size caps how many of each can exist. The header counts
	// are untrusted, so clamp to this bound to keep a corrupt file (e.g. a
	// huge numDIFATSects with a cyclic DIFAT chain) from driving an
	// effectively unbounded amount of work.
	maxSects := len(data) / f.sectorSize

	numFATSects := binary.LittleEndian.Uint32(data[44:])
	firstDirSect := binary.LittleEndian.Uint32(data[48:])
	firstMiniFATSect := binary.LittleEndian.Uint32(data[60:])
	numMiniFATSects := binary.LittleEndian.Uint32(data[64:])
	firstDIFATSect := binary.LittleEndian.Uint32(data[68:])
	numDIFATSects := binary.LittleEndian.Uint32(data[72:])

	// DIFAT: 109 entries in the header, then a chain of DIFAT sectors
	// ([MS-CFB] 2.5). Each DIFAT sector ends with the next sector number.
	var fatSects []uint32
	for i := 0; i < 109; i++ {
		s := binary.LittleEndian.Uint32(data[76+i*4:])
		if s <= secMaxRegular {
			fatSects = append(fatSects, s)
		}
	}
	difatSect := firstDIFATSect
	entriesPerDIFAT := f.sectorSize/4 - 1
	if numDIFATSects > uint32(maxSects) {
		numDIFATSects = uint32(maxSects)
	}
	for i := uint32(0); i < numDIFATSects && difatSect <= secMaxRegular; i++ {
		sect, err := f.sectorData(difatSect)
		if err != nil {
			return nil, fmt.Errorf("DIFAT: %w", err)
		}
		for j := 0; j < entriesPerDIFAT; j++ {
			s := binary.LittleEndian.Uint32(sect[j*4:])
			if s <= secMaxRegular {
				fatSects = append(fatSects, s)
			}
		}
		difatSect = binary.LittleEndian.Uint32(sect[entriesPerDIFAT*4:])
	}
	if numFATSects > uint32(maxSects) {
		numFATSects = uint32(maxSects)
	}
	if uint32(len(fatSects)) < numFATSects {
		// Tolerate header miscounts; use what we found.
		numFATSects = uint32(len(fatSects))
	}

	// FAT ([MS-CFB] 2.4): concatenation of all FAT sectors.
	for _, s := range fatSects[:numFATSects] {
		sect, err := f.sectorData(s)
		if err != nil {
			return nil, fmt.Errorf("FAT: %w", err)
		}
		for j := 0; j+4 <= len(sect); j += 4 {
			f.fat = append(f.fat, binary.LittleEndian.Uint32(sect[j:]))
		}
	}

	// Directory ([MS-CFB] 2.6): chain starting at firstDirSect.
	dirData, err := f.readChain(firstDirSect, 0)
	if err != nil {
		return nil, fmt.Errorf("directory: %w", err)
	}
	for off := 0; off+128 <= len(dirData); off += 128 {
		e := dirData[off : off+128]
		nameLen := int(binary.LittleEndian.Uint16(e[64:]))
		if nameLen < 2 || nameLen > 64 {
			continue
		}
		u := make([]uint16, (nameLen-2)/2)
		for i := range u {
			u[i] = binary.LittleEndian.Uint16(e[i*2:])
		}
		f.dirs = append(f.dirs, dirEntry{
			name:      string(utf16.Decode(u)),
			objType:   e[66],
			startSect: binary.LittleEndian.Uint32(e[116:]),
			size:      binary.LittleEndian.Uint64(e[120:]),
		})
	}
	if len(f.dirs) == 0 || f.dirs[0].objType != 5 {
		return nil, errors.New("missing root directory entry")
	}
	if major == 3 {
		// [MS-CFB] 2.6.1: in version 3 files only the low 32 bits of the
		// stream size are meaningful.
		for i := range f.dirs {
			f.dirs[i].size &= 0xFFFFFFFF
		}
	}

	// Mini FAT ([MS-CFB] 2.8) and the mini stream (root entry's chain).
	if firstMiniFATSect <= secMaxRegular && numMiniFATSects > 0 {
		miniFATData, err := f.readChain(firstMiniFATSect, 0)
		if err != nil {
			return nil, fmt.Errorf("mini FAT: %w", err)
		}
		for j := 0; j+4 <= len(miniFATData); j += 4 {
			f.miniFAT = append(f.miniFAT, binary.LittleEndian.Uint32(miniFATData[j:]))
		}
		f.miniStream, err = f.readChain(f.dirs[0].startSect, f.dirs[0].size)
		if err != nil {
			return nil, fmt.Errorf("mini stream: %w", err)
		}
	}
	return f, nil
}

// sectorData returns the raw bytes of one regular sector.
func (f *cfbFile) sectorData(sect uint32) ([]byte, error) {
	off := (int64(sect) + 1) * int64(f.sectorSize)
	if off < 0 || off+int64(f.sectorSize) > int64(len(f.data)) {
		return nil, fmt.Errorf("sector %d out of range", sect)
	}
	return f.data[off : off+int64(f.sectorSize)], nil
}

// readChain follows a FAT sector chain. size==0 reads the whole chain.
func (f *cfbFile) readChain(start uint32, size uint64) ([]byte, error) {
	var out []byte
	sect := start
	for steps := 0; sect <= secMaxRegular; steps++ {
		if steps > len(f.fat) {
			return nil, errors.New("FAT chain loop detected")
		}
		d, err := f.sectorData(sect)
		if err != nil {
			return nil, err
		}
		out = append(out, d...)
		if sect >= uint32(len(f.fat)) {
			return nil, fmt.Errorf("sector %d beyond FAT", sect)
		}
		sect = f.fat[sect]
	}
	if size > 0 {
		if uint64(len(out)) < size {
			return nil, fmt.Errorf("stream truncated: want %d bytes, chain has %d", size, len(out))
		}
		out = out[:size]
	}
	return out, nil
}

// readMiniChain follows a mini FAT chain inside the mini stream
// (64-byte mini sectors, [MS-CFB] 2.8).
func (f *cfbFile) readMiniChain(start uint32, size uint64) ([]byte, error) {
	const miniSize = 64
	var out []byte
	sect := start
	for steps := 0; sect <= secMaxRegular; steps++ {
		if steps > len(f.miniFAT) {
			return nil, errors.New("mini FAT chain loop detected")
		}
		off := int(sect) * miniSize
		if off+miniSize > len(f.miniStream) {
			return nil, fmt.Errorf("mini sector %d out of range", sect)
		}
		out = append(out, f.miniStream[off:off+miniSize]...)
		if sect >= uint32(len(f.miniFAT)) {
			return nil, fmt.Errorf("mini sector %d beyond mini FAT", sect)
		}
		sect = f.miniFAT[sect]
	}
	if uint64(len(out)) < size {
		return nil, fmt.Errorf("mini stream truncated: want %d bytes, chain has %d", size, len(out))
	}
	return out[:size], nil
}

// openStream returns the contents of the named stream anywhere in the
// directory. Streams smaller than the mini stream cutoff (normally 4096
// bytes) live in the mini stream ([MS-CFB] 2.6.1).
func (f *cfbFile) openStream(name string) ([]byte, error) {
	for _, e := range f.dirs {
		if e.objType == 2 && e.name == name {
			if e.size == 0 {
				return nil, nil
			}
			if e.size < uint64(f.miniCutoff) {
				return f.readMiniChain(e.startSect, e.size)
			}
			return f.readChain(e.startSect, e.size)
		}
	}
	return nil, fmt.Errorf("stream %q not found", name)
}

// hasStream reports whether a stream with the given name exists.
func (f *cfbFile) hasStream(name string) bool {
	for _, e := range f.dirs {
		if e.objType == 2 && e.name == name {
			return true
		}
	}
	return false
}

// streamNames lists the names of all streams in the file.
func (f *cfbFile) streamNames() []string {
	var names []string
	for _, e := range f.dirs {
		if e.objType == 2 {
			names = append(names, e.name)
		}
	}
	return names
}
