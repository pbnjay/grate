package crypto

import (
	"bytes"
	"crypto/md5"
	"crypto/rc4"
	"encoding/binary"
	"fmt"
)

var _ Decryptor = &rc4Writer{}

func (d *rc4Writer) Write(data []byte) (n int, err error) {
	x := len(data)
	for len(data) > 0 {
		n := copy(d.bytes[d.offset:], data)
		d.offset += n
		if d.offset >= 1024 {
			if d.offset != 1024 {
				panic("invalid offset from write")
			}
			d.Flush()
		}
		data = data[n:]
	}
	return x, nil
}

func (d *rc4Writer) Read(data []byte) (n int, err error) {
	return d.buf.Read(data)
}

// Reset to block 0, and clear all written and readable data.
func (d *rc4Writer) Reset() {
	d.block = 0
	d.offset = 0
	d.buf.Reset()
}

// Flush tells the decryptor to decrypt the latest block.
func (d *rc4Writer) Flush() {
	var zeros [1024]byte

	endpad := 0
	if d.offset < 1024 {
		endpad = copy(d.bytes[d.offset:], zeros[:])
		d.offset += endpad
	}
	if d.offset != 1024 {
		panic("invalid offset fill")
	}

	// decrypt and write results to output buffer
	d.startBlock()
	d.dec.XORKeyStream(d.bytes[:], d.bytes[:])
	d.buf.Write(d.bytes[:1024-endpad])

	d.offset = 0
	d.block++
}

// SetPassword for the decryption.
func (d *rc4Writer) SetPassword(password []byte) {
	d.Password = make([]rune, len(password))
	for i, p := range password {
		d.Password[i] = rune(p)
	}

	/// compute the first part of the encryption key
	result := generateStd97Key(d.Password, d.Salt)
	d.encKey = make([]byte, len(result))
	copy(d.encKey, result)
}

type rc4Writer struct {
	block  uint32
	offset int
	bytes  [1024]byte

	// records the decrypted data
	buf bytes.Buffer

	///////

	// decrypter for RC4 content streams
	dec *rc4.Cipher

	cipherKey []byte // H1 per 2.3.6.2
	encKey    []byte // Hfinal per 2.3.6.2

	Salt     []byte
	Password []rune
}

func (d *rc4Writer) Verify(everifier, everifierHash []byte) error {
	d.Reset()
	d.startBlock()

	var temp1 [16]byte
	var temp2 [16]byte
	d.dec.XORKeyStream(temp1[:], everifier)
	d.dec.XORKeyStream(temp2[:], everifierHash)

	newhash := md5.Sum(temp1[:])
	for i, c := range newhash {
		if temp2[i] != c {
			return fmt.Errorf("verification failed")
		}
	}
	return nil
}

/////////////////////

func (d *rc4Writer) startBlock() {
	if d.encKey == nil {
		d.SetPassword([]byte(DefaultXLSPassword))
	}

	d.cipherKey = make([]byte, 16)
	copy(d.cipherKey, d.encKey[:5])
	binary.LittleEndian.PutUint32(d.cipherKey[5:], d.block)
	mhash := md5.Sum(d.cipherKey[:9])
	d.dec, _ = rc4.NewCipher(mhash[:])
}

func generateStd97Key(passData []rune, salt []byte) []byte {
	if len(passData) == 0 || len(salt) != 16 {
		panic("invalid keygen material")
	}

	passBytes := make([]byte, len(passData)*2)

	for i, c := range passData {
		binary.LittleEndian.PutUint16(passBytes[2*i:], uint16(c))
	}

	// digest the IV then copy back into pKeyData
	h0 := md5.Sum(passBytes)

	// now do the final set of keygen ops
	msum := md5.New()
	for i := 0; i < 16; i++ {
		msum.Write(h0[:5])
		msum.Write(salt)
	}
	// return H1
	temp := make([]byte, 0, 16)
	temp = msum.Sum(temp)
	return temp
}
