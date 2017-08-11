package doc2txt

import (
	"testing"
)

func TestGetFib(t *testing.T) {
	_, err := getFib(nil)
	if err != errDocEmpty {
		t.Error("Expected error due to empty WordDoc")
	}
	if _, err := getFib(reader.File[2]); err == nil { // short mscfb.File
		t.Error("expected error due to short file", err)
	}
	if _, err = getFib(reader.File[4]); err != errFibInvalid { // use wrong mscfb.File
		t.Error("expected error due to corrupt file", err)
	}

	fib, _ := getFib(simpleDoc)
	if fib.csw != 28 || fib.cslw != 88 || fib.cbRgFcLcb != 0x00B7 {
		t.Error("expected valid sizes", fib.csw, fib.cslw, fib.cbRgFcLcb)
	}
	if fib.base.fWhichTblStm != 1 {
		t.Error("expected table 1")
	}
	// No headers in simpleDoc, just "12345" in the text which apparently makes a ccpText of 6
	// cpLength is calculated and should equal ccpText in this scenario
	if fib.fibRgLw.ccpAtn != 0 || fib.fibRgLw.ccpEdn != 0 || fib.fibRgLw.ccpFtn != 0 || fib.fibRgLw.ccpHdd != 0 || fib.fibRgLw.ccpHdrTxbx != 0 ||
		fib.fibRgLw.ccpMcr != 0 || fib.fibRgLw.ccpText != 6 || fib.fibRgLw.cpLength != 6 {
		t.Error("expected valid fibRgLw", fib.fibRgLw)
	}
	// These are the values in the byte stream at the correct locations
	if fib.fibRgFcLcb.fcClx != 5279 || fib.fibRgFcLcb.lcbClx != 21 {
		t.Error("expected valid fibRgFcLcb", fib.fibRgFcLcb)
	}
}
