# doc2txt
[![Build Status](https://travis-ci.org/EndFirstCorp/doc2txt.svg?branch=master)](https://travis-ci.org/EndFirstCorp/doc2txt) [![Coverage Status](https://coveralls.io/repos/github/EndFirstCorp/doc2txt/badge.svg?branch=master)](https://coveralls.io/github/EndFirstCorp/doc2txt?branch=master)

A native Go reader for the old Microsoft Word .doc binary format files

Example usage:

   f, _ := os.Open(`testData\simpleDoc.doc`)
   buf, err := ParseDoc(f)
   if err != nil {
     // handle error
   }
   // buf now contains an io.Reader which you can save to the file system or further transform

## Special Thanks
A great big thank you to Richard Lehane. His [(https://github.com/richardlehane/mscfb](https://github.com/richardlehane/mscfb) got me started, his [https://github.com/richardlehane/doctool](https://github.com/richardlehane/doctool) project got me closer and his answer to questions via email helped get me to the finish line. Thanks Richard!