package doc2txt

import (
	"encoding/binary"
	"errors"

	"github.com/richardlehane/mscfb"
)

var (
	errFibInvalid = errors.New("file information block validation failed")
)

type fib struct {
	base       fibBase
	csw        int
	fibRgW     fibRgW
	cslw       int
	fibRgLw    fibRgLw
	cbRgFcLcb  int
	fibRgFcLcb fibRgFcLcb
}

type fibBase struct {
	fWhichTblStm int
}

type fibRgW struct {
}

type fibRgLw struct {
	ccpText    int
	ccpFtn     int
	ccpHdd     int
	ccpMcr     int
	ccpAtn     int
	ccpEdn     int
	ccpTxbx    int
	ccpHdrTxbx int
	cpLength   int
}

type fibRgFcLcb struct {
	fcClx  int
	lcbClx int
}

// parse File Information Block (section 2.5.1)
func getFib(wordDoc *mscfb.File) (*fib, error) {
	if wordDoc == nil {
		return nil, errDocEmpty
	}

	b := make([]byte, 894) // get FIB block up to FibRgFcLcb97
	_, err := wordDoc.ReadAt(b, 0)
	if err != nil {
		return nil, err
	}

	fibBase := getFibBase(b[0:32])

	fibRgW, csw, err := getFibRgW(b, 32)
	if err != nil {
		return nil, err
	}

	fibRgLw, cslw, err := getFibRgLw(b, 34+csw)
	if err != nil {
		return nil, err
	}

	fibRgFcLcb, cbRgFcLcb, err := getFibRgFcLcb(b, 34+csw+2+cslw)

	return &fib{base: *fibBase, csw: csw, cslw: cslw, fibRgW: *fibRgW, fibRgLw: *fibRgLw, fibRgFcLcb: *fibRgFcLcb, cbRgFcLcb: cbRgFcLcb}, err
}

// parse FibBase (section 2.5.2)
func getFibBase(fib []byte) *fibBase {
	byt := fib[11]                    // fWhichTblStm is 2nd highest bit in this byte
	fWhichTblStm := int(byt >> 1 & 1) // set which table (0Table or 1Table) is the table stream
	return &fibBase{fWhichTblStm: fWhichTblStm}
}

func getFibRgW(fib []byte, start int) (*fibRgW, int, error) {
	if start+2 >= len(fib) { // must be big enough for csw
		return &fibRgW{}, 0, errFibInvalid
	}

	csw := int(binary.LittleEndian.Uint16(fib[start:start+2])) * 2 // in bytes
	return &fibRgW{}, csw, nil
}

// parse FibRgLw (section 2.5.4)
func getFibRgLw(fib []byte, start int) (*fibRgLw, int, error) {
	fibRgLwStart := start + 2 // skip cslw
	ccpTextLoc := fibRgLwStart + 3*4
	ccpFtnLoc := fibRgLwStart + 4*4
	ccpHddLoc := fibRgLwStart + 5*4
	ccpMcrLoc := fibRgLwStart + 6*4
	ccpAtnLoc := fibRgLwStart + 7*4
	ccpEdnLoc := fibRgLwStart + 8*4
	ccpTxbxLoc := fibRgLwStart + 9*4
	ccpHdrTxbxLoc := fibRgLwStart + 10*4
	if ccpHdrTxbxLoc+4 >= len(fib) { // must be big enough for ccpHdrTxbx
		return &fibRgLw{}, 0, errFibInvalid
	}

	cslw := int(binary.LittleEndian.Uint16(fib[start:start+2])) * 4 // in bytes
	ccpText := int(binary.LittleEndian.Uint32(fib[ccpTextLoc : ccpTextLoc+4]))
	ccpFtn := int(binary.LittleEndian.Uint32(fib[ccpFtnLoc : ccpFtnLoc+4]))
	ccpHdd := int(binary.LittleEndian.Uint32(fib[ccpHddLoc : ccpHddLoc+4]))
	ccpMcr := int(binary.LittleEndian.Uint32(fib[ccpMcrLoc : ccpMcrLoc+4]))
	ccpAtn := int(binary.LittleEndian.Uint32(fib[ccpAtnLoc : ccpAtnLoc+4]))
	ccpEdn := int(binary.LittleEndian.Uint32(fib[ccpEdnLoc : ccpEdnLoc+4]))
	ccpTxbx := int(binary.LittleEndian.Uint32(fib[ccpTxbxLoc : ccpTxbxLoc+4]))
	ccpHdrTxbx := int(binary.LittleEndian.Uint32(fib[ccpHdrTxbxLoc : ccpHdrTxbxLoc+4]))

	// calculate cpLength. Used in PlcPcd verification (see section 2.8.35)
	var cpLength int
	if ccpFtn != 0 || ccpHdd != 0 || ccpMcr != 0 || ccpAtn != 0 || ccpEdn != 0 || ccpTxbx != 0 || ccpHdrTxbx != 0 {
		cpLength = ccpFtn + ccpHdd + ccpMcr + ccpAtn + ccpEdn + ccpEdn + ccpTxbx + ccpHdrTxbx + ccpText + 1
	} else {
		cpLength = ccpText
	}
	return &fibRgLw{ccpText: ccpText, ccpFtn: ccpFtn, ccpHdd: ccpHdd, ccpMcr: ccpMcr, ccpAtn: ccpAtn,
		ccpEdn: ccpEdn, ccpTxbx: ccpTxbx, ccpHdrTxbx: ccpHdrTxbx, cpLength: cpLength}, cslw, nil
}

// parse FibRgFcLcb (section 2.5.5)
func getFibRgFcLcb(fib []byte, start int) (*fibRgFcLcb, int, error) {
	fibRgFcLcbStart := start + 2 // skip cbRgFcLcb
	fcClxLoc := fibRgFcLcbStart + 66*4
	lsbClxLoc := fibRgFcLcbStart + 67*4
	if lsbClxLoc+4 >= len(fib) { // must be big enough for lsbClxLoc
		return &fibRgFcLcb{}, 0, errFibInvalid
	}

	cbRgFcLcb := int(binary.LittleEndian.Uint16(fib[start : start+2]))
	fcClx := int(binary.LittleEndian.Uint32(fib[fcClxLoc : fcClxLoc+4]))
	lcbClx := int(binary.LittleEndian.Uint32(fib[lsbClxLoc : lsbClxLoc+4]))
	return &fibRgFcLcb{fcClx: fcClx, lcbClx: lcbClx}, cbRgFcLcb, nil
}
