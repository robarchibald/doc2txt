package doc2txt

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"io"

	"github.com/mattetti/filebuffer"
	"github.com/richardlehane/mscfb"
)

var (
	errNoFields    = errors.New("No fields")
	errFibShort    = errors.New("file information block too short")
	errClxShort    = errors.New("clx block too short")
	errTable       = errors.New("cannot find table stream")
	errInvalidClx  = errors.New("expected last aCP value to equal fib.cpLength (2.8.35)")
	errDocShort    = errors.New("wordDoc block too short")
	errInvalidPcdt = errors.New("expected clxt to be equal 0x02")
)

type allReader interface {
	io.Closer
	io.ReaderAt
	io.ReadSeeker
}

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

type clx struct {
	pcdt pcdt
}

type pcdt struct {
	lcb    int
	PlcPcd plcPcd
}

type plcPcd struct {
	aCP  []int
	aPcd []pcd
}

type pcd struct {
	fc fcCompressed
}

type fcCompressed struct {
	fc          int
	fCompressed bool
}

func wrapError(e error) error {
	return errors.New("Error processing file: " + e.Error())
}

// ParseDoc converts a standard io.Reader from a Microsoft Word
// .doc binary file and returns a reader (actually a bytes.Buffer)
// which will output the plain text found in the .doc file
func ParseDoc(r io.Reader) (io.Reader, error) {
	ra, ok := r.(io.ReaderAt)
	if !ok {
		ra, _, err := toMemoryBuffer(r)
		if err != nil {
			return nil, wrapError(err)
		}
		defer ra.Close()
	}

	d, err := mscfb.New(ra)
	if err != nil {
		return nil, wrapError(err)
	}

	wordDoc, table0, table1 := getWordDocAndTables(d)
	fib, err := getFib(wordDoc)
	if err != nil {
		return nil, wrapError(err)
	}

	table := getActiveTable(table0, table1, fib)
	if table == nil {
		return nil, wrapError(errTable)
	}

	clx, err := getClx(table, fib)
	if err != nil {
		return nil, wrapError(err)
	}
	if clx.pcdt.PlcPcd.aCP[len(clx.pcdt.PlcPcd.aCP)-1] != fib.fibRgLw.cpLength {
		return nil, wrapError(errInvalidClx)
	}

	fmt.Println(clx)
	fmt.Println(getText(wordDoc, clx))

	return &bytes.Buffer{}, nil
}

func toMemoryBuffer(r io.Reader) (allReader, int64, error) {
	var b bytes.Buffer
	size, err := b.ReadFrom(r)
	if err != nil {
		return nil, 0, err
	}
	fb := filebuffer.New(b.Bytes())
	return fb, size, nil
}

func getText(wordDoc *mscfb.File, clx *clx) (string, error) {
	var buf bytes.Buffer
	for i := 0; i < len(clx.pcdt.PlcPcd.aPcd); i++ {
		pcd := clx.pcdt.PlcPcd.aPcd[i]
		cp := clx.pcdt.PlcPcd.aCP[i]
		cpNext := clx.pcdt.PlcPcd.aCP[i+1]

		var start, end, size int
		if pcd.fc.fCompressed {
			size = 1
			start = pcd.fc.fc / 2
			end = start + cpNext - cp
		} else {
			size = 2
			start = pcd.fc.fc
			end = start + 2*(cpNext-cp)
		}

		b := make([]byte, end-start)
		i, err := wordDoc.ReadAt(b, int64(start/size))
		if err != nil {
			return "", err
		} else if i != end-start {
			return "", errDocShort
		}
		buf.Write(b)
	}
	return buf.String(), nil
}

// read Clx (section 2.9.38)
func getClx(table *mscfb.File, fib *fib) (*clx, error) {
	b, err := readClx(table, fib)
	if err != nil {
		return nil, err
	}

	pcdtOffset := getPrcArrayEnd(b)
	pcdt, err := getPcdt(b, pcdtOffset)
	if err != nil {
		return nil, err
	}
	return &clx{pcdt: *pcdt}, nil
}

func readClx(table *mscfb.File, fib *fib) ([]byte, error) {
	b := make([]byte, fib.fibRgFcLcb.lcbClx)
	i, err := table.ReadAt(b, int64(fib.fibRgFcLcb.fcClx))
	if err != nil {
		return nil, err
	}
	if i != fib.fibRgFcLcb.lcbClx {
		return nil, errClxShort // clx is not long enough
	}
	return b, nil
}

// read Pcdt from Clx (section )
func getPcdt(clx []byte, pcdtOffset int) (*pcdt, error) {
	const pcdSize = 8
	if clx[pcdtOffset] != 0x02 { // clxt must be 0x02 or invalid
		return nil, errInvalidPcdt
	}
	lcb := int(binary.LittleEndian.Uint32(clx[pcdtOffset+1 : pcdtOffset+5])) // skip clxt, get lcb
	plcPcdOffset := pcdtOffset + 5                                           // skip clxt and lcb
	numPcds := (lcb - 4) / (4 + pcdSize)                                     // see 2.2.2 in the spec for equation
	numCps := numPcds + 1                                                    // always 1 more cp than pcds

	cps := make([]int, numCps)
	for i := 0; i < numCps; i++ {
		cpOffset := plcPcdOffset + i*4
		cps[i] = int(binary.LittleEndian.Uint32(clx[cpOffset : cpOffset+4]))
	}

	pcdStart := plcPcdOffset + 4*numCps
	pcds := make([]pcd, numPcds)
	for i := 0; i < numPcds; i++ {
		pcdOffset := pcdStart + i*pcdSize
		pcds[i] = *parsePcd(clx[pcdOffset : pcdOffset+pcdSize])
	}
	return &pcdt{lcb: lcb, PlcPcd: plcPcd{aCP: cps, aPcd: pcds}}, nil
}

// find end of RgPrc array (section 2.9.38)
func getPrcArrayEnd(clx []byte) int {
	prcOffset := 0
	for {
		clxt := clx[prcOffset]
		if clxt != 0x01 { // this is not a Prc, so exit
			return prcOffset
		}
		prcDataCbGrpprl, _ := binary.Varint(clx[prcOffset+1 : prcOffset+3]) // skip the clxt and read 2 bytes
		prcOffset = prcOffset + 1 + 2 + int(prcDataCbGrpprl)                // skip clxt, cbGrpprl, and GrpPrl
	}
}

// parse Pcd (section 2.9.177)
func parsePcd(pcdData []byte) *pcd {
	return &pcd{fc: *parseFcCompressed(pcdData[2:6])}
}

// parse FcCompressed (section 2.9.73)
func parseFcCompressed(fcData []byte) *fcCompressed {
	data := binary.BigEndian.Uint32(fcData) // we're reading in bits, so use big endian
	fCompressed := (data >> 1 & 1) == 1     // get rid of all but bit 2
	fc := data >> 2                         // shift out the lowest order bits

	return &fcCompressed{fc: int(fc), fCompressed: fCompressed}
}

// parse File Information Block (section 2.5.1)
func getFib(wordDoc *mscfb.File) (*fib, error) {
	if wordDoc == nil {
		return nil, errors.New("WordDocument not found")
	}

	b := make([]byte, 894) // get FIB block up to FibRgFcLcb97
	i, err := wordDoc.Read(b)
	if err != nil {
		return nil, err
	}
	if i < 894 {
		return nil, errFibShort // fib is not long enough
	}

	fibBase := getFibBase(b[0:32])

	csw := int(binary.LittleEndian.Uint16(b[32:34])) * 2 // in bytes
	fibRgwStart := 34
	//fibRgW := getFibRgW(b[fibRgwStart : fibRgwStart+csw])

	cslw := int(binary.LittleEndian.Uint16(b[fibRgwStart+csw:fibRgwStart+csw+2])) * 4 // in bytes
	fibRgLwStart := fibRgwStart + csw + 2
	fibRgLw := getFibRgLw(b[fibRgLwStart : fibRgLwStart+cslw*4])

	cbRgFcLcb := int(binary.LittleEndian.Uint16(b[fibRgLwStart+cslw : fibRgLwStart+cslw+2]))
	fibRgFcLcbStart := fibRgLwStart + cslw + 2
	fibRgFcLcb := getFibRgFcLcb(b[fibRgFcLcbStart : fibRgFcLcbStart+cbRgFcLcb])

	return &fib{base: *fibBase, csw: csw, cslw: cslw, fibRgLw: *fibRgLw, fibRgFcLcb: *fibRgFcLcb, cbRgFcLcb: cbRgFcLcb}, nil
}

// parse FibBase (section 2.5.2)
func getFibBase(fib []byte) *fibBase {
	byt := fib[11]                    // fWhichTblStm is 2nd highest bit in this byte
	fWhichTblStm := int(byt >> 1 & 1) // set which table (0Table or 1Table) is the table stream
	return &fibBase{fWhichTblStm: fWhichTblStm}
}

// parse FibRgLw (section 2.5.4)
func getFibRgLw(fib []byte) *fibRgLw {
	const ccpTextLoc = 3 * 4
	const ccpFtnLoc = 4 * 4
	const ccpHddLoc = 5 * 4
	const ccpMcrLoc = 6 * 4
	const ccpAtnLoc = 7 * 4
	const ccpEdnLoc = 8 * 4
	const ccpTxbxLoc = 9 * 4
	const ccpHdrTxbxLoc = 10 * 4
	ccpText := int(binary.LittleEndian.Uint32(fib[ccpTextLoc : ccpTextLoc+4]))
	ccpFtn := int(binary.LittleEndian.Uint32(fib[ccpFtnLoc : ccpFtnLoc+4]))
	ccpHdd := int(binary.LittleEndian.Uint32(fib[ccpHddLoc : ccpHddLoc+4]))
	ccpMcr := int(binary.LittleEndian.Uint32(fib[ccpMcrLoc : ccpMcrLoc+4]))
	ccpAtn := int(binary.LittleEndian.Uint32(fib[ccpAtnLoc : ccpAtnLoc+4]))
	ccpEdn := int(binary.LittleEndian.Uint32(fib[ccpEdnLoc : ccpEdnLoc+4]))
	ccpTxbx := int(binary.LittleEndian.Uint32(fib[ccpTxbxLoc : ccpTxbxLoc+4]))
	ccpHdrTxbx := int(binary.LittleEndian.Uint32(fib[ccpHdrTxbxLoc : ccpHdrTxbxLoc+4]))

	var cpLength int
	if ccpFtn != 0 || ccpHdd != 0 || ccpMcr != 0 || ccpAtn != 0 || ccpEdn != 0 || ccpTxbx != 0 || ccpHdrTxbx != 0 {
		cpLength = ccpFtn + ccpHdd + ccpMcr + ccpAtn + ccpEdn + ccpEdn + ccpTxbx + ccpHdrTxbx + ccpText + 1
	} else {
		cpLength = ccpText
	}
	return &fibRgLw{ccpText: ccpText, ccpFtn: ccpFtn, ccpHdd: ccpHdd, ccpMcr: ccpMcr, ccpAtn: ccpAtn,
		ccpEdn: ccpEdn, ccpTxbx: ccpTxbx, ccpHdrTxbx: ccpHdrTxbx, cpLength: cpLength}
}

// parse FibRgFcLcb (section 2.5.5)
func getFibRgFcLcb(fib []byte) *fibRgFcLcb {
	const fcClxLoc = 66 * 4
	const lsbClxLoc = 67 * 4
	fcClx := int(binary.LittleEndian.Uint32(fib[fcClxLoc : fcClxLoc+4]))
	lcbClx := int(binary.LittleEndian.Uint32(fib[lsbClxLoc : lsbClxLoc+4]))
	return &fibRgFcLcb{fcClx: fcClx, lcbClx: lcbClx}
}

func getWordDocAndTables(r *mscfb.Reader) (*mscfb.File, *mscfb.File, *mscfb.File) {
	var wordDoc, table0, table1 *mscfb.File
	for i := 0; i < len(r.File); i++ {
		stream := r.File[i]

		switch stream.Name {
		case "WordDocument":
			wordDoc = stream
		case "0Table":
			table0 = stream
		case "1Table":
			table1 = stream
		}
	}
	return wordDoc, table0, table1
}

func getActiveTable(table0 *mscfb.File, table1 *mscfb.File, f *fib) *mscfb.File {
	if f.base.fWhichTblStm == 0 {
		return table0
	}
	return table1
}
