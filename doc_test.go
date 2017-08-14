package doc2txt

import (
	"bytes"
	"os"
	"strings"
	"testing"
)

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

	var expected bytes.Buffer
	expected.WriteString(strings.Replace(complicatedDoc, "\n", "\r", -1))

	actual := buf.(*bytes.Buffer)
	count := 0
	for aline, err := actual.ReadString('\r'); err == nil; aline, err = actual.ReadString('\r') {
		eline, err := expected.ReadString('\r')
		if err != nil || eline != aline {
			t.Errorf("mismatch at line %d. Expected: %s, Actual: %s\n", count, eline, aline)
		}
		count++
	}
}

const complicatedDoc = `Name Here in Big
Link to something



Summary
Testing out new things

Bullet 1
Bullet 2
Bullet 3

Underlined
Italics

Numbered list
Item 1
Item 2
Item 3

Some Information  In a Table  Hopefully, we get it  


Here is some information with a footnote
Here is some information with an endnote

Here is a table of contents

Contents
Some	1
Information	1
In a	1
Table	1
Hopefully, we	1
get it	1
	1



Header 1
Header 2
Header 3

 Here is my footnote









Information in the header


Some Footer information	current date:8/7/2017 4:16:33 PM	pg. 1


 My endnote

Some info from inside a text box



`
