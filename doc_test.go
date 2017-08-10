package doc2txt

import (
	"os"
	"testing"
)

func TestParseDoc(t *testing.T) {
	f, err := os.Open(`testData\simpleDoc.doc`)
	_, err = ParseDoc(f)
	if err != nil {
		t.Error(err)
	}

}
