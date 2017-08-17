using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

public class FileStream_ThatDeletesFileAfterReading : FileStream
{
    public FileStream_ThatDeletesFileAfterReading(string path, FileMode mode, FileAccess access)
        : base(path, mode, access)
    {
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        if (File.Exists(Name))
            File.Delete(Name);
    }
}