package crypto

import (
	"bytes"
	"encoding/binary"
	"fmt"
)

// Decryptor describes methods to decrypt an excel sheet.
type Decryptor interface {
	// SetPassword for the decryption.
	SetPassword(password []byte)

	// Read implements the io.Reader interface.
	Read(p []byte) (n int, err error)

	// Write implements the io.Writer interface.
	Write(p []byte) (n int, err error)

	// Flush tells the decryptor to decrypt the latest block.
	Flush()

	// Reset the decryptor, and clear all written and readable data.
	Reset()
}

// Algorithms designed based on specs in MS-OFFCRYPTO:
// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/3c34d72a-1a61-4b52-a893-196f9157f083

// Important notes from MS-XLS section 2.2.10:
// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/cd03cb5f-ca02-4934-a391-bb674cb8aa06

// When obfuscating or encrypting BIFF records in these streams the record type and
// record size components MUST NOT be obfuscated or encrypted.
// In addition the following records MUST NOT be obfuscated or encrypted:
// BOF (section 2.4.21), FilePass (section 2.4.117), UsrExcl (section 2.4.339),
// FileLock (section 2.4.116), InterfaceHdr (section 2.4.146), RRDInfo (section 2.4.227),
// and RRDHead (section 2.4.226). Additionally, the lbPlyPos field of the BoundSheet8
// record (section 2.4.28) MUST NOT be encrypted.

// For RC4 encryption and RC4 CryptoAPI encryption, the Unicode password string is used
// to generate the encryption key as specified in [MS-OFFCRYPTO] section 2.3.6.2 or
// [MS-OFFCRYPTO] section 2.3.5.2 depending on the RC4 algorithm used. The record data
// is then encrypted by the specific RC4 algorithm in 1024-byte blocks. The block number
// is set to zero at the beginning of every BIFF record stream, and incremented by one
// at each 1024-byte boundary. Bytes to be encrypted are passed into the RC4 encryption
// function and then written to the stream. For unencrypted records and the record
// headers consisting of the record type and record size, a byte buffer of all zeros,
// of the same size as the section of unencrypted bytes, is passed into the RC4
// encryption function. The results are then ignored and the unencrypted bytes are
// written to the stream.

// DefaultXLSPassword is the default encryption password defined by note
// <100> Section 2.4.191: If the value of the wPassword field of the Password record in
// the Globals Substream is not 0x0000, Excel 97, Excel 2000, Excel 2002, Office Excel
// 2003, Office Excel 2007, and Excel 2010 encrypt the document as specified in [MS-OFFCRYPTO],
// section 2.3. If an encryption password is not specified or the workbook or sheet is only
// protected, the document is encrypted with the default password of:
var DefaultXLSPassword = "VelvetSweatshop"

/////////////

// 2.3.6.1
type basicRC4Encryption struct {
	MajorVersion uint16
	MinorVersion uint16
	Salt         [16]byte
	Verifier     [16]byte
	VerifierHash [16]byte
}

// NewBasicRC4 implements the standard RC4 decryption.
func NewBasicRC4(data []byte) (Decryptor, error) {
	h := basicRC4Encryption{}
	b := bytes.NewReader(data)
	err := binary.Read(b, binary.LittleEndian, &h)
	if err != nil {
		return nil, err
	}
	if h.MinorVersion != 1 {
		return nil, fmt.Errorf("xls: unknown basic-RC4 minor version %d (%d byte record)",
			h.MinorVersion, len(data))
	}
	if len(data) != 52 {
		return nil, fmt.Errorf("xls: data length is invalid (expected 52 bytes, got %d)",
			len(data))
	}

	d := &rc4Writer{
		Salt: make([]byte, len(h.Salt)),
	}
	copy(d.Salt, h.Salt[:])

	return d, d.Verify(h.Verifier[:], h.VerifierHash[:])
}
