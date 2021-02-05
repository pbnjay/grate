package xls

type header struct {
	Version  uint16 // An unsigned integer that specifies the BIFF version of the file. The value MUST be 0x0600.
	DocType  uint16 //An unsigned integer that specifies the document type of the substream of records following this record. For more information about the layout of the sub-streams in the workbook stream see File Structure.
	RupBuild uint16 // An unsigned integer that specifies the build identifier.
	RupYear  uint16 // An unsigned integer that specifies the year when this BIFF version was first created. The value MUST be 0x07CC or 0x07CD.
	MiscBits uint64 // lots of miscellaneous bits and flags we're not going to check
}

// 2.1.4
type rec struct {
	RecType recordType //
	RecSize uint16     // must be between 0 and 8224
	Data    []byte     // len(rec.data) = rec.recsize
}

type boundSheet struct {
	Position    uint32 // A FilePointer as specified in [MS-OSHARED] section 2.2.1.5 that specifies the stream position of the start of the BOF record for the sheet.
	HiddenState byte   // (2 bits) An unsigned integer that specifies the hidden state of the sheet. MUST be a value from the following table:
	SheetType   byte   // An unsigned integer that specifies the sheet type. 00=worksheet
	Name        string
}
