package doc2txt

import (
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

func TestParseDoc(t *testing.T) {
	f, err := os.Open(`testData\simpleDoc.doc`)
	_, err = ParseDoc(f)
	if err != nil {
		t.Error(err)
	}

}
