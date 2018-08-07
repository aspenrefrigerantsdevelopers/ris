<%

Function Encrypt(value)
    dim EncryptedValue, objEncrypt
    set objEncrypt = server.CreateObject("CUUtils.Encrypt")
    
    EncryptedValue = objEncrypt.EncryptStringAscii(cstr(value))
    set objEncrypt = Nothing
    
    Encrypt = EncryptedValue

End Function 


Function Decrypt(value)
    dim DecryptedValue, objEncrypt
    set objEncrypt = server.CreateObject("CUUtils.Encrypt")
    DecryptedValue = objEncrypt.EncryptStringAscii(cstr(value))
    set objEncrypt = Nothing
    
    Decrypt = DecryptedValue

End Function 


 %>