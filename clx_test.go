package doc2txt

import (
	"testing"
)

func TestGetClx(t *testing.T) {
	// invalid argument(s)
	_, err := getClx(nil, nil)
	if err != errInvalidArgument {
		t.Error("expected invalid argument", err)
	}
	// can't read empty table
	f := &fib{fibRgFcLcb: fibRgFcLcb{fcClx: 12, lcbClx: 21}}
	if _, err = getClx(reader.File[0], f); err == nil {
		t.Error("expected error reading")
	}
	// invalid read location, invalid data
	if _, err = getClx(table, f); err != errInvalidPcdt {
		t.Error("expected error reading", err)
	}
	// all correct, but cpLength
	f = &fib{fibRgFcLcb: fibRgFcLcb{fcClx: 5279, lcbClx: 21}}
	if _, err = getClx(table, f); err != errInvalidClx {
		t.Error("expected error reading", err)
	}

	// values come from a successful parse of simpleDoc.doc
	f.fibRgLw = fibRgLw{cpLength: 6}
	clx, err := getClx(table, f)
	if err != nil || clx.pcdt.lcb != 16 || len(clx.pcdt.PlcPcd.aCP) != 2 || len(clx.pcdt.PlcPcd.aPcd) != 1 ||
		clx.pcdt.PlcPcd.aCP[0] != 0 || clx.pcdt.PlcPcd.aCP[1] != 6 ||
		clx.pcdt.PlcPcd.aPcd[0].fc.fc != 4096 || clx.pcdt.PlcPcd.aPcd[0].fc.fCompressed != true {
		t.Error("expected valid clx", clx, err)
	}
}

func TestGetPrcArrayEnd(t *testing.T) {
	//skip since it is not a Prc
	clx := []byte{2, 0, 0, 0}
	if num, _ := getPrcArrayEnd(clx); num != 0 {
		t.Error("expected to be set to beginning")
	}
	// error due to zero offset with valid Prc clxt
	clx = []byte{1, 0, 0, 0}
	if _, err := getPrcArrayEnd(clx); err != errInvalidPrc {
		t.Error("expected to revert to 0 due to invalid value", err)
	}
	// error since next offset would be too large
	clx = []byte{1, 4, 4, 0}
	if _, err := getPrcArrayEnd(clx); err != errInvalidPrc {
		t.Error("expected to revert to 0 due to invalid value", err)
	}
	// two items
	clx = []byte{1, 2, 0, 0, 0, 1, 2, 0, 0, 2, 2, 2, 2}
	if num, err := getPrcArrayEnd(clx); err != nil || num != 10 {
		t.Error("expected to revert to 0 due to invalid value", err, num)
	}
}
