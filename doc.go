package doc2txt

import (
	"bytes"
	"errors"
	"io"

	"github.com/mattetti/filebuffer"
	"github.com/richardlehane/mscfb"
)

var (
	errTable           = errors.New("cannot find table stream")
	errDocEmpty        = errors.New("WordDocument not found")
	errDocShort        = errors.New("wordDoc block too short")
	errInvalidArgument = errors.New("invalid table and/or fib")
)

type allReader interface {
	io.Closer
	io.ReaderAt
	io.ReadSeeker
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

	return getText(wordDoc, clx)
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

func getText(wordDoc *mscfb.File, clx *clx) (io.Reader, error) {
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
			return nil, err
		} else if i != end-start {
			return nil, errDocShort
		}
		buf.Write(b)
	}
	return &buf, nil
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
