package doc2txt

import (
	"bytes"
	"fmt"
	"os"
	"testing"

	"github.com/richardlehane/mscfb"
)

var simpleDoc *mscfb.File
var table *mscfb.File
var reader *mscfb.Reader

func init() {
	f, _ := os.Open(`testData\simpleDoc.doc`)
	reader, _ = mscfb.New(f)
	simpleDoc, _, table = getWordDocAndTables(reader)
}

func TestParseSimpleDoc(t *testing.T) {
	f, _ := os.Open(`testData\simpleDoc.doc`)
	buf, err := ParseDoc(f)
	if err != nil {
		t.Fatal("expected successful parse", err)
	}
	if s := buf.(*bytes.Buffer).String(); s != "12345\r" {
		t.Errorf("expected correct value |%s|", s)
	}
}

func TestParseComplicated(t *testing.T) {
	f, _ := os.Open(`testData\docFile.doc`)
	buf, err := ParseDoc(f)
	if err != nil {
		t.Fatal("expected to be able to parse document", err)
	}
	if err != nil || buf.(*bytes.Buffer).String() != "hello" {
		fmt.Println(buf.(*bytes.Buffer).String())
		t.Error("expected to be able to parse document", err)
	}
}
