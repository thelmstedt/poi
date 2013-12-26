/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package org.apache.poi.poifs.crypt.agile;

import static org.apache.poi.poifs.crypt.CryptoFunctions.generateIv;
import static org.apache.poi.poifs.crypt.CryptoFunctions.getBlock0;
import static org.apache.poi.poifs.crypt.CryptoFunctions.getCipher;
import static org.apache.poi.poifs.crypt.CryptoFunctions.getMessageDigest;
import static org.apache.poi.poifs.crypt.CryptoFunctions.hashPassword;
import static org.apache.poi.poifs.crypt.agile.AgileDecryptor.getNextBlockSize;
import static org.apache.poi.poifs.crypt.agile.AgileDecryptor.hashInput;
import static org.apache.poi.poifs.crypt.agile.AgileDecryptor.kCryptoKeyBlock;
import static org.apache.poi.poifs.crypt.agile.AgileDecryptor.kHashedVerifierBlock;
import static org.apache.poi.poifs.crypt.agile.AgileDecryptor.kIntegrityKeyBlock;
import static org.apache.poi.poifs.crypt.agile.AgileDecryptor.kIntegrityValueBlock;
import static org.apache.poi.poifs.crypt.agile.AgileDecryptor.kVerifierInputBlock;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilterOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.security.GeneralSecurityException;
import java.security.MessageDigest;
import java.security.SecureRandom;
import java.security.cert.CertificateEncodingException;
import java.security.spec.AlgorithmParameterSpec;
import java.util.HashMap;
import java.util.Map;
import java.util.Random;

import javax.crypto.Cipher;
import javax.crypto.Mac;
import javax.crypto.SecretKey;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.RC2ParameterSpec;
import javax.crypto.spec.SecretKeySpec;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.poifs.crypt.CipherAlgorithm;
import org.apache.poi.poifs.crypt.CryptoFunctions;
import org.apache.poi.poifs.crypt.DataSpaceMapUtils;
import org.apache.poi.poifs.crypt.EncryptionHeader;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.apache.poi.poifs.crypt.agile.AgileEncryptionVerifier.AgileCertificateEntry;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.POIFSWriterEvent;
import org.apache.poi.poifs.filesystem.POIFSWriterListener;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.util.LittleEndianByteArrayOutputStream;
import org.apache.poi.util.LittleEndianConsts;
import org.apache.poi.util.LittleEndianOutputStream;
import org.apache.poi.util.TempFile;
import org.apache.xmlbeans.XmlOptions;

import com.microsoft.schemas.office.x2006.encryption.CTDataIntegrity;
import com.microsoft.schemas.office.x2006.encryption.CTEncryption;
import com.microsoft.schemas.office.x2006.encryption.CTKeyData;
import com.microsoft.schemas.office.x2006.encryption.CTKeyEncryptor;
import com.microsoft.schemas.office.x2006.encryption.CTKeyEncryptors;
import com.microsoft.schemas.office.x2006.encryption.EncryptionDocument;
import com.microsoft.schemas.office.x2006.encryption.STCipherAlgorithm;
import com.microsoft.schemas.office.x2006.encryption.STCipherChaining;
import com.microsoft.schemas.office.x2006.encryption.STHashAlgorithm;
import com.microsoft.schemas.office.x2006.keyEncryptor.certificate.CTCertificateKeyEncryptor;
import com.microsoft.schemas.office.x2006.keyEncryptor.password.CTPasswordKeyEncryptor;

public class AgileEncryptor extends Encryptor {
    private final AgileEncryptionInfoBuilder builder;
    @SuppressWarnings("unused")
    private byte integritySalt[];
    private Mac integrityMD;
	private byte pwHash[];
    
	protected AgileEncryptor(AgileEncryptionInfoBuilder builder) {
		this.builder = builder;
	}

    public void confirmPassword(String password) {
        // see [MS-OFFCRYPTO] - 2.3.3 EncryptionVerifier
        Random r = new SecureRandom();
        int blockSize = builder.getHeader().getBlockSize();
        int keySize = builder.getHeader().getKeySize()/8;
        int hashSize = builder.getHeader().getHashAlgorithmEx().hashSize;
        
        byte[] verifierSalt = new byte[blockSize]
             , verifier = new byte[blockSize]
             , keySalt = new byte[blockSize]
             , keySpec = new byte[keySize]
             , integritySalt = new byte[hashSize];
        r.nextBytes(verifierSalt); // blocksize
        r.nextBytes(verifier); // blocksize
        r.nextBytes(keySalt); // blocksize
        r.nextBytes(keySpec); // keysize
        r.nextBytes(integritySalt); // hashsize
        
        confirmPassword(password, keySpec, keySalt, verifierSalt, verifier, integritySalt);
    }
	
	public void confirmPassword(String password, byte keySpec[], byte keySalt[], byte verifier[], byte verifierSalt[], byte integritySalt[]) {
        AgileEncryptionVerifier ver = builder.getVerifier();
        ver.setSalt(verifierSalt);
        AgileEncryptionHeader header = builder.getHeader();
        header.setKeySalt(keySalt);
        HashAlgorithm hashAlgo = ver.getHashAlgorithm();

        int blockSize = header.getBlockSize();
	    
        pwHash = hashPassword(password, hashAlgo, verifierSalt, ver.getSpinCount());
        
        /**
         * encryptedVerifierHashInput: This attribute MUST be generated by using the following steps:
         * 1. Generate a random array of bytes with the number of bytes used specified by the saltSize
         *    attribute.
         * 2. Generate an encryption key as specified in section 2.3.4.11 by using the user-supplied password,
         *    the binary byte array used to create the saltValue attribute, and a blockKey byte array
         *    consisting of the following bytes: 0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, and 0x79.
         * 3. Encrypt the random array of bytes generated in step 1 by using the binary form of the saltValue
         *    attribute as an initialization vector as specified in section 2.3.4.12. If the array of bytes is not an
         *    integral multiple of blockSize bytes, pad the array with 0x00 to the next integral multiple of
         *    blockSize bytes.
         * 4. Use base64 to encode the result of step 3.
         */
        byte encryptedVerifier[] = hashInput(builder, pwHash, kVerifierInputBlock, verifier, Cipher.ENCRYPT_MODE);
        ver.setEncryptedVerifier(encryptedVerifier);
	    

        /**
         * encryptedVerifierHashValue: This attribute MUST be generated by using the following steps:
         * 1. Obtain the hash value of the random array of bytes generated in step 1 of the steps for
         *    encryptedVerifierHashInput.
         * 2. Generate an encryption key as specified in section 2.3.4.11 by using the user-supplied password,
         *    the binary byte array used to create the saltValue attribute, and a blockKey byte array
         *    consisting of the following bytes: 0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, and 0x4e.
         * 3. Encrypt the hash value obtained in step 1 by using the binary form of the saltValue attribute as
         *    an initialization vector as specified in section 2.3.4.12. If hashSize is not an integral multiple of
         *    blockSize bytes, pad the hash value with 0x00 to an integral multiple of blockSize bytes.
         * 4. Use base64 to encode the result of step 3.
         */
        MessageDigest hashMD = getMessageDigest(hashAlgo);
        byte[] hashedVerifier = hashMD.digest(verifier);
        byte encryptedVerifierHash[] = hashInput(builder, pwHash, kHashedVerifierBlock, hashedVerifier, Cipher.ENCRYPT_MODE);
        ver.setEncryptedVerifierHash(encryptedVerifierHash);
        
        /**
         * encryptedKeyValue: This attribute MUST be generated by using the following steps:
         * 1. Generate a random array of bytes that is the same size as specified by the
         *    Encryptor.KeyData.keyBits attribute of the parent element.
         * 2. Generate an encryption key as specified in section 2.3.4.11, using the user-supplied password,
         *    the binary byte array used to create the saltValue attribute, and a blockKey byte array
         *    consisting of the following bytes: 0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, and 0xd6.
         * 3. Encrypt the random array of bytes generated in step 1 by using the binary form of the saltValue
         *    attribute as an initialization vector as specified in section 2.3.4.12. If the array of bytes is not an
         *    integral multiple of blockSize bytes, pad the array with 0x00 to an integral multiple of
         *    blockSize bytes.
         * 4. Use base64 to encode the result of step 3.
         */
        byte encryptedKey[] = hashInput(builder, pwHash, kCryptoKeyBlock, keySpec, Cipher.ENCRYPT_MODE);
        ver.setEncryptedKey(encryptedKey);
        
        SecretKey secretKey = new SecretKeySpec(keySpec, ver.getCipherAlgorithm().jceId);
        setSecretKey(secretKey);
        
        /*
         * 2.3.4.14 DataIntegrity Generation (Agile Encryption)
         * 
         * The DataIntegrity element contained within an Encryption element MUST be generated by using
         * the following steps:
         * 1. Obtain the intermediate key by decrypting the encryptedKeyValue from a KeyEncryptor
         *    contained within the KeyEncryptors sequence. Use this key for encryption operations in the
         *    remaining steps of this section.
         * 2. Generate a random array of bytes, known as Salt, of the same length as the value of the
         *    KeyData.hashSize attribute.
         * 3. Encrypt the random array of bytes generated in step 2 by using the binary form of the
         *    KeyData.saltValue attribute and a blockKey byte array consisting of the following bytes:
         *    0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, and 0xf6 used to form an initialization vector as
         *    specified in section 2.3.4.12. If the array of bytes is not an integral multiple of blockSize
         *    bytes, pad the array with 0x00 to the next integral multiple of blockSize bytes.
         * 4. Assign the encryptedHmacKey attribute to the base64-encoded form of the result of step 3.
         * 5. Generate an HMAC, as specified in [RFC2104], of the encrypted form of the data (message),
         *    which the DataIntegrity element will verify by using the Salt generated in step 2 as the key.
         *    Note that the entire EncryptedPackage stream (1), including the StreamSize field, MUST be
         *    used as the message.
         * 6. Encrypt the HMAC as in step 3 by using a blockKey byte array consisting of the following bytes:
         *    0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, and 0x33.
         * 7.  Assign the encryptedHmacValue attribute to the base64-encoded form of the result of step 6. 
         */
        this.integritySalt = integritySalt;

        try {
            byte vec[] = CryptoFunctions.generateIv(hashAlgo, header.getKeySalt(), kIntegrityKeyBlock, header.getBlockSize());
            Cipher cipher = getCipher(secretKey, ver.getCipherAlgorithm(), ver.getChainingMode(), vec, Cipher.ENCRYPT_MODE);
            byte filledSalt[] = getBlock0(integritySalt, getNextBlockSize(integritySalt.length, blockSize));
            byte encryptedHmacKey[] = cipher.doFinal(filledSalt);
            header.setEncryptedHmacKey(encryptedHmacKey);

            this.integrityMD = CryptoFunctions.getMac(hashAlgo);
            this.integrityMD.init(new SecretKeySpec(integritySalt, hashAlgo.jceHmacId));

        
            cipher = Cipher.getInstance("RSA");
            for (AgileCertificateEntry ace : ver.getCertificates()) {
                cipher.init(Cipher.ENCRYPT_MODE, ace.x509.getPublicKey());
                ace.encryptedKey = cipher.doFinal(getSecretKey().getEncoded());
                Mac x509Hmac = CryptoFunctions.getMac(hashAlgo);
                x509Hmac.init(getSecretKey());
                ace.certVerifier = x509Hmac.doFinal(ace.x509.getEncoded());
            }
        } catch (GeneralSecurityException e) {
            throw new EncryptedDocumentException(e);
        }
	}
	
    public OutputStream getDataStream(DirectoryNode dir)
            throws IOException, GeneralSecurityException {
        // TODO: initialize headers
        OutputStream countStream = new ChunkedCipherOutputStream(dir);
    	return countStream;
    }

    /**
     * 2.3.4.15 Data Encryption (Agile Encryption)
     * 
     * The EncryptedPackage stream (1) MUST be encrypted in 4096-byte segments to facilitate nearly
     * random access while allowing CBC modes to be used in the encryption process.
     * The initialization vector for the encryption process MUST be obtained by using the zero-based
     * segment number as a blockKey and the binary form of the KeyData.saltValue as specified in
     * section 2.3.4.12. The block number MUST be represented as a 32-bit unsigned integer.
     * Data blocks MUST then be encrypted by using the initialization vector and the intermediate key
     * obtained by decrypting the encryptedKeyValue from a KeyEncryptor contained within the
     * KeyEncryptors sequence as specified in section 2.3.4.10. The final data block MUST be padded to
     * the next integral multiple of the KeyData.blockSize value. Any padding bytes can be used. Note
     * that the StreamSize field of the EncryptedPackage field specifies the number of bytes of
     * unencrypted data as specified in section 2.3.4.4.
     */
    private class ChunkedCipherOutputStream extends FilterOutputStream implements POIFSWriterListener {
        private long _pos = 0;
        private final byte[] _chunk = new byte[4096];
        private Cipher _cipher;
        private final File fileOut;
        protected final DirectoryNode dir;

        public ChunkedCipherOutputStream(DirectoryNode dir) throws IOException {
            super(null);
            fileOut = TempFile.createTempFile("encrypted_package", "crypt");
            this.out = new FileOutputStream(fileOut);
            this.dir = dir;
            EncryptionHeader header = builder.getHeader();
            _cipher = getCipher(getSecretKey(), header.getCipherAlgorithm(), header.getChainingMode(), null, Cipher.ENCRYPT_MODE);
        }

        public void write(int b) throws IOException {
            write(new byte[]{(byte)b});
        }

        public void write(byte[] b) throws IOException {
            write(b, 0, b.length);
        }

        public void write(byte[] b, int off, int len)
        throws IOException {
            if (len == 0) return;
            
            if (len < 0 || b.length < off+len) {
                throw new IOException("not enough bytes in your input buffer");
            }
            
            while (len > 0) {
                int posInChunk = (int)(_pos & 0xfff);
                int nextLen = Math.min(4096-posInChunk, len);
                System.arraycopy(b, off, _chunk, posInChunk, nextLen);
                _pos += nextLen;
                off += nextLen;
                len -= nextLen;
                if ((_pos & 0xfff) == 0) {
                    writeChunk();
                }
            }
        }

        private void writeChunk() throws IOException {
            EncryptionHeader header = builder.getHeader();
            int blockSize = header.getBlockSize();

            int posInChunk = (int)(_pos & 0xfff);
            // normally posInChunk is 0, i.e. on the next chunk (-> index-1)
            // but if called on close(), posInChunk is somewhere within the chunk data
            int index = (int)(_pos >> 12);
            if (posInChunk==0) {
                index--;
                posInChunk = 4096;
            } else {
                // pad the last chunk
                _cipher = getCipher(getSecretKey(), header.getCipherAlgorithm(), header.getChainingMode(), null, Cipher.ENCRYPT_MODE, "PKCS5Padding");
            }

            byte[] blockKey = new byte[4];
            LittleEndian.putInt(blockKey, 0, index);
            byte[] iv = generateIv(header.getHashAlgorithmEx(), header.getKeySalt(), blockKey, blockSize);
            try {
                AlgorithmParameterSpec aps;
                if (header.getCipherAlgorithm() == CipherAlgorithm.rc2) {
                    aps = new RC2ParameterSpec(getSecretKey().getEncoded().length*8, iv);
                } else {
                    aps = new IvParameterSpec(iv);
                }
                
                _cipher.init(Cipher.ENCRYPT_MODE, getSecretKey(), aps);
                int ciLen = _cipher.doFinal(_chunk, 0, posInChunk, _chunk);
                out.write(_chunk, 0, ciLen);
            } catch (GeneralSecurityException e) {
                throw (IOException)new IOException().initCause(e);
            }
        }
        
        public void close() throws IOException {
            writeChunk();
            super.close();
            writeToPOIFS();
        }

        void writeToPOIFS() throws IOException {
            DataSpaceMapUtils.addDefaultDataSpace(dir);
            
            /**
             * Generate an HMAC, as specified in [RFC2104], of the encrypted form of the data (message), 
             * which the DataIntegrity element will verify by using the Salt generated in step 2 as the key. 
             * Note that the entire EncryptedPackage stream (1), including the StreamSize field, MUST be 
             * used as the message.
             * 
             * Encrypt the HMAC as in step 3 by using a blockKey byte array consisting of the following bytes:
             * 0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, and 0x33.
             **/
            byte buf[] = new byte[4096];
            LittleEndian.putLong(buf, 0, _pos);
            integrityMD.update(buf, 0, LittleEndianConsts.LONG_SIZE);
            
            InputStream fis = new FileInputStream(fileOut);
            for (int readBytes; (readBytes = fis.read(buf)) != -1; integrityMD.update(buf, 0, readBytes));
            fis.close();
            
            AgileEncryptionHeader header = builder.getHeader(); 
            int blockSize = header.getBlockSize();
            
            byte hmacValue[] = integrityMD.doFinal();
            byte iv[] = CryptoFunctions.generateIv(header.getHashAlgorithmEx(), header.getKeySalt(), kIntegrityValueBlock, header.getBlockSize());
            Cipher cipher = CryptoFunctions.getCipher(getSecretKey(), header.getCipherAlgorithm(), header.getChainingMode(), iv, Cipher.ENCRYPT_MODE);
            try {
                byte hmacValueFilled[] = getBlock0(hmacValue, getNextBlockSize(hmacValue.length, blockSize));
                byte encryptedHmacValue[] = cipher.doFinal(hmacValueFilled);
                header.setEncryptedHmacValue(encryptedHmacValue);
            } catch (GeneralSecurityException e) {
                throw new EncryptedDocumentException(e);
            }

            createEncryptionInfoEntry(dir);
            
            int oleStreamSize = (int)(fileOut.length()+LittleEndianConsts.LONG_SIZE);
            dir.createDocument("EncryptedPackage", oleStreamSize, this);
            // TODO: any properties???
        }
    
        public void processPOIFSWriterEvent(POIFSWriterEvent event) {
            try {
                LittleEndianOutputStream leos = new LittleEndianOutputStream(event.getStream());

                // StreamSize (8 bytes): An unsigned integer that specifies the number of bytes used by data 
                // encrypted within the EncryptedData field, not including the size of the StreamSize field. 
                // Note that the actual size of the \EncryptedPackage stream (1) can be larger than this 
                // value, depending on the block size of the chosen encryption algorithm
                leos.writeLong(_pos);

                FileInputStream fis = new FileInputStream(fileOut);
                IOUtils.copy(fis, leos);
                fis.close();
                fileOut.delete();

                leos.close();
            } catch (IOException e) {
                throw new EncryptedDocumentException(e);
            }
        }
    }

    protected void createEncryptionInfoEntry(DirectoryNode dir) throws IOException {
        AgileEncryptionVerifier ver = builder.getVerifier();
        AgileEncryptionHeader header = builder.getHeader();
        
        EncryptionDocument ed = EncryptionDocument.Factory.newInstance();
        CTEncryption edRoot = ed.addNewEncryption();
        
        CTKeyData keyData = edRoot.addNewKeyData();
        CTKeyEncryptors keyEncList = edRoot.addNewKeyEncryptors();
        CTKeyEncryptor keyEnc = keyEncList.addNewKeyEncryptor();
        keyEnc.setUri(CTKeyEncryptor.Uri.HTTP_SCHEMAS_MICROSOFT_COM_OFFICE_2006_KEY_ENCRYPTOR_PASSWORD);
        CTPasswordKeyEncryptor keyPass = keyEnc.addNewEncryptedPasswordKey();

        keyPass.setSpinCount(ver.getSpinCount());
        
        keyData.setSaltSize(header.getBlockSize());
        keyPass.setSaltSize(header.getBlockSize());
        
        keyData.setBlockSize(header.getBlockSize());
        keyPass.setBlockSize(header.getBlockSize());

        keyData.setKeyBits(header.getKeySize());
        keyPass.setKeyBits(header.getKeySize());

        HashAlgorithm hashAlgo = header.getHashAlgorithmEx();
        keyData.setHashSize(hashAlgo.hashSize);
        keyPass.setHashSize(hashAlgo.hashSize);

        STCipherAlgorithm.Enum xmlCipherAlgo = STCipherAlgorithm.Enum.forString(header.getCipherAlgorithm().xmlId);
        if (xmlCipherAlgo == null) {
            throw new EncryptedDocumentException("CipherAlgorithm "+header.getCipherAlgorithm()+" not supported.");
        }
        keyData.setCipherAlgorithm(xmlCipherAlgo);
        keyPass.setCipherAlgorithm(xmlCipherAlgo);
        
        switch (header.getChainingMode()) {
        case cbc: 
            keyData.setCipherChaining(STCipherChaining.CHAINING_MODE_CBC);
            keyPass.setCipherChaining(STCipherChaining.CHAINING_MODE_CBC);
            break;
        case cfb:
            keyData.setCipherChaining(STCipherChaining.CHAINING_MODE_CFB);
            keyPass.setCipherChaining(STCipherChaining.CHAINING_MODE_CFB);
            break;
        default:
            throw new EncryptedDocumentException("ChainingMode "+header.getChainingMode()+" not supported.");
        }
        
        STHashAlgorithm.Enum xmlHashAlgo = STHashAlgorithm.Enum.forString(hashAlgo.ecmaString);
        if (xmlHashAlgo == null) {
            throw new EncryptedDocumentException("HashAlgorithm "+hashAlgo+" not supported.");
        }
        keyData.setHashAlgorithm(xmlHashAlgo);
        keyPass.setHashAlgorithm(xmlHashAlgo);

        keyData.setSaltValue(header.getKeySalt());
        keyPass.setSaltValue(ver.getSalt());
        keyPass.setEncryptedVerifierHashInput(ver.getEncryptedVerifier());
        keyPass.setEncryptedVerifierHashValue(ver.getEncryptedVerifierHash());
        keyPass.setEncryptedKeyValue(ver.getEncryptedKey());
        
        CTDataIntegrity hmacData = edRoot.addNewDataIntegrity();
        hmacData.setEncryptedHmacKey(header.getEncryptedHmacKey());
        hmacData.setEncryptedHmacValue(header.getEncryptedHmacValue());
        
        for (AgileCertificateEntry ace : ver.getCertificates()) {
            keyEnc = keyEncList.addNewKeyEncryptor();
            keyEnc.setUri(CTKeyEncryptor.Uri.HTTP_SCHEMAS_MICROSOFT_COM_OFFICE_2006_KEY_ENCRYPTOR_CERTIFICATE);
            CTCertificateKeyEncryptor certData = keyEnc.addNewEncryptedCertificateKey();
            try {
                certData.setX509Certificate(ace.x509.getEncoded());
            } catch (CertificateEncodingException e) {
                throw new EncryptedDocumentException(e);
            }
            certData.setEncryptedKeyValue(ace.encryptedKey);
            certData.setCertVerifier(ace.certVerifier);
        }

        XmlOptions xo = new XmlOptions();
        xo.setCharacterEncoding("UTF-8");
        Map<String,String> nsMap = new HashMap<String,String>();
        nsMap.put("http://schemas.microsoft.com/office/2006/keyEncryptor/password","p");
        nsMap.put("http://schemas.microsoft.com/office/2006/keyEncryptor/certificate", "c");
        nsMap.put("http://schemas.microsoft.com/office/2006/encryption","");
        xo.setSaveSuggestedPrefixes(nsMap);
        xo.setSaveNamespacesFirst();
        xo.setSaveAggressiveNamespaces();
        // setting standalone doesn't work with xmlbeans-2.3
        xo.setSaveNoXmlDecl();
        
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        bos.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n".getBytes("UTF-8"));
        ed.save(bos, xo);

        final byte buf[] = new byte[5000];        
        LittleEndianByteArrayOutputStream leos = new LittleEndianByteArrayOutputStream(buf, 0);
        EncryptionInfo info = builder.getInfo();

        // EncryptionVersionInfo (4 bytes): A Version structure (section 2.1.4), where 
        // Version.vMajor MUST be 0x0004 and Version.vMinor MUST be 0x0004
        leos.writeShort(info.getVersionMajor());
        leos.writeShort(info.getVersionMinor());
        // Reserved (4 bytes): A value that MUST be 0x00000040
        leos.writeInt(0x40);
        leos.write(bos.toByteArray());
        
        dir.createDocument("EncryptionInfo", leos.getWriteIndex(), new POIFSWriterListener() {
            public void processPOIFSWriterEvent(POIFSWriterEvent event) {
                try {
                    event.getStream().write(buf, 0, event.getLimit());
                } catch (IOException e) {
                    throw new EncryptedDocumentException(e);
                }
            }
        });
    }
}