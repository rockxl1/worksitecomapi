using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XTRWorkSite.WSObjects.Enums
{
    public enum XTR_EnumFolderAttributeId
    {
        imFolderID = 0,
        imFolderParentID = 1,
        imFolderDefaultSecurity = 2,
        imFolderName = 3,
        imFolderOwner = 4,
        imFolderDescription = 5,
        imFolderLocation = 6,
        imFolderType = 7,
        imFolderSubtype = 8,
        imFolderInherits = 9,
        imFolderProfileID = 10,
        imFolderVersion = 11,
        //imFolderCustom1 = 12, //API TEM UM BUG. não pesquisa por customs!
        //imFolderCustom2 = 13,
        //imFolderCustom3 = 14,
        imFolderRootFolderID = 15,
        imFolderRootFolderType = 16,
        imFolderIsContentSearch = 17,
        imFolderIsFolderSearch = 18,
        imFolderEmail = 19,
        imFolderHiddenOnDesktop = 20,
        imFolderEditDate = 21,
        imFolderReferenceType = 22,
    }
}
