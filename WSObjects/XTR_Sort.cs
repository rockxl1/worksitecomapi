using IManage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XTRWorkSite.WSObjects
{
    public class XTR_Sort : IManObjectSort
    {
        /// <summary>
        /// Member variable
        /// </summary>
        Field m_Field;

        /// <summary>
        /// Constructor: Creates an instance of this class for sorting items in a collection.
        /// </summary>
        public XTR_Sort()
        {
            // Default sort field value;
            //	if user does not select from Fiedl enums, sort by Client-matter.
            m_Field = Field.Client_matter;
        }

        /// <summary>
        /// Enum SortByField: 
        /// define a sample of sort fields for user from which to select.
        /// </summary>
        public enum Field : int
        {
            Owner,
            Description,
            Name,
            Client_matter
        }

        /// <summary>
        /// This property is used to set the field user has selected by which to sort.
        /// </summary>
        public Field FieldToSort
        {
            set { m_Field = value; }
        }

        /// <summary>
        /// This method is used by this class to query for the folder attribute value
        ///	matching the custom field name in 'fieldName'.  Returns the string value if
        ///	found.  If not, or if errored, returns an empty string.
        /// </summary>
        /// <param name="props">The additional properties collection</param>
        /// <param name="fieldName">The target fieldname to check for in the additional properties collection</param>
        /// <returns>The string value of the additional property</returns>
        private string getAttributeValue(IManAdditionalProperties props, string fieldName)
        {
            if (props.Count > 0)
            {
                try
                {
                    if (props.ContainsByName(fieldName) == true)
                        return props.ItemByName(fieldName).Value;
                }
                catch (Exception) { }
            }
            return string.Empty;
        }

        /// <summary>
        /// This method is used by this class to compare two strings
        /// under the same field to determine whether str1 is less than str2.
        /// </summary>
        /// <param name="str1">The string value to be compared</param>
        /// <param name="str2">The string value to be compared</param>
        /// <returns>A boolean value indicating whether str1 is less than str2</returns>
        private bool SingleFieldCompare(string str1, string str2)
        {
            // Initialize result to false
            bool result = false;

            if ((populated(str1) == true) &&
                (populated(str2) == true))
            {
                if (str1.CompareTo(str2) < 0)
                    // First item is less.
                    result = true;
            }
            if ((populated(str1) == false) &&
                (populated(str2) == true))
            {
                // First item is less.
                result = true;
            }
            return result;
        }

        public bool populated(string content)
        {
            try
            {
                return (content != null) && (content.Length > 0);
            }
            catch (Exception e)
            {
            }
            return false;
        }

        /// <summary>
        /// This method is used by this class to:
        ///	1. compare two strings under the same parent field
        ///		to determine whether strParent1 is less than strParent2.
        ///	2. if the two parent strings are equal, compare 
        ///		the two strings	under the same child field to further
        ///		determine whether strParent1 is less than strParent2.
        /// </summary>
        /// <param name="strParent1">The string value in the parent field to be compared</param>
        /// <param name="strParent2">The string value in the parent field to be compared</param>
        /// <param name="strChild1">The string value in the child field to be compared</param>
        /// <param name="strChild2">The string value in the child field to be compared</param>
        /// <returns>A boolean that indicates that the first set of parent/child field value is less than the second set</returns>
        private bool DoubleFieldCompare(string strParent1, string strParent2, string strChild1, string strChild2)
        {
            // Initialize result to false
            bool result = false;

            if ((populated(strParent1) == true) &&
                (populated(strParent2) == true))
            // Both parent strings are not empty.
            {
                if ((strParent1.CompareTo(strParent2)) < 0)
                    // First item is less.
                    result = true;
                if ((strParent1.CompareTo(strParent2)) == 0)
                // Both parent strings are equal.
                {
                    if ((populated(strChild1) == true) &&
                        (populated(strChild2) == true))
                    // Both child strings are not empty.
                    {
                        if ((strChild1.CompareTo(strChild2)) < 0)
                            // First item is less
                            result = true;
                    }
                    if ((populated(strChild1) == false) &&
                        (populated(strChild2) == true))
                    {
                        // First item is less.
                        result = true;
                    }
                }
            }
            else if ((populated(strParent1) == false) &&
                (populated(strParent2) == true))
            {
                // First item is less.
                result = true;
            }
            return result;
        }

        #region IManObjectSort Members

        /// <summary>
        /// This method implements the code logic to determine if object1 is less than objec2.  
        /// If so, returns true; otherwise, returns false.  The two objects in comparison
        /// will get sorted in ascending order if the return value is true.
        /// </summary>
        /// <param name="object1">The object to be compared</param>
        /// <param name="object2">The object to be compared</param>
        /// <returns>A boolean that indicates that the first object is less than the second</returns>
        public bool Less(IManObject object1, IManObject object2)
        {      
            try
            {
                IManDocument m_Doc1 = (IManDocument)object1;
                IManDocument m_Doc2 = (IManDocument)object2;

                if(m_Doc1.Number > m_Doc2.Number)
                {
                    return true;
                }
                else
                {
                    return false;
                }

                // Initialize return value to False.
                bool retVal = false;

                /*
				 * NOTE: the naming convention for querying a document
				 * profile field value from AdditionalProperties.ItemByName
				 * Method of a folder object is as follows:
				 * prefix of "iman___" appending to the imProfileAttributeID Enum.
				 * Example: "iman___imProfileCustom1"
				*/

                // Variable for doc profile field prefix.
                string strPrefix = "iman___";

                // Variables for doc profile fields Custom1 & Custom2.
                string strCust1 = strPrefix + "imProfileCustom1";
                string strCust2 = strPrefix + "imProfileCustom2";

                // Variable to hold base object type
                IManObjectType objType = object1.ObjectType;

                // Identify the strongly typed object.
                // NOTE: The following does not comprise the entire
                //			set of collection object types that can be sorted.
                switch (objType.ObjectType)
                {
                    // Find out the object type then sort
                    case imObjectType.imTypeDocument:
                        // It's document
                       

                        // Sort by the field user has selected.
                        if (Field.Description == m_Field)
                        {
                            // Compare the two document descriptons.
                            retVal = this.SingleFieldCompare(m_Doc1.Description, m_Doc2.Description);
                        }
                        else if (Field.Name == m_Field)
                        {
                            // Compare the two document names.
                            retVal = this.SingleFieldCompare(m_Doc1.Name, m_Doc2.Name);
                        }
                        else if (Field.Owner == m_Field)
                        {
                            // Compare the two document authors.
                            retVal = this.SingleFieldCompare(m_Doc1.Author.Name, m_Doc2.Author.Name);
                        }
                        else if (Field.Client_matter == m_Field) // default
                        {
                            // Compare the parent and child values of both documents.		
                            retVal = this.DoubleFieldCompare
                                        (m_Doc1.GetAttributeValueByID(imProfileAttributeID.imProfileCustom1).ToString(),
                                            m_Doc2.GetAttributeValueByID(imProfileAttributeID.imProfileCustom1).ToString(),
                                            m_Doc1.GetAttributeValueByID(imProfileAttributeID.imProfileCustom2).ToString(),
                                            m_Doc2.GetAttributeValueByID(imProfileAttributeID.imProfileCustom2).ToString());
                        }
                     
                        break;
                    case imObjectType.imTypeFavoritesFolder:
                        // It's favorite folder
                        IManFavoritesFolder m_FavFldr1 = (IManFavoritesFolder)object1;
                        IManFavoritesFolder m_FavFldr2 = (IManFavoritesFolder)object2;

                        // Sort by the field user has selected.
                        if (Field.Description == m_Field)
                        {
                            // Compare the two folder descriptons.
                            retVal = this.SingleFieldCompare(m_FavFldr1.Description, m_FavFldr2.Description);
                        }
                        else if (Field.Name == m_Field)
                        {
                            // Compare the two folder names.
                            retVal = this.SingleFieldCompare(m_FavFldr1.Name, m_FavFldr2.Name);
                        }
                        else if (Field.Owner == m_Field)
                        {
                            // Compare the two folder owner names.
                            retVal = this.SingleFieldCompare(m_FavFldr1.Owner.Name, m_FavFldr2.Owner.Name);
                        }
                        else if (Field.Client_matter == m_Field) // default
                        {
                            // Query for profile attributes Custom1 & Custom2, and
                            //	compare the parent and child values of both folders.		
                            retVal = this.DoubleFieldCompare(
                                            (this.getAttributeValue(m_FavFldr1.AdditionalProperties, strCust1)),
                                            (this.getAttributeValue(m_FavFldr2.AdditionalProperties, strCust1)),
                                            (this.getAttributeValue(m_FavFldr1.AdditionalProperties, strCust2)),
                                            (this.getAttributeValue(m_FavFldr2.AdditionalProperties, strCust2)));
                        }
                        break;
                    case imObjectType.imTypeDocumentFolder:
                        // It's document folder
                        IManDocumentFolder m_DocFldr1 = (IManDocumentFolder)object1;
                        IManDocumentFolder m_DocFldr2 = (IManDocumentFolder)object2;

                        // Sort by the field user has selected.
                        if (Field.Description == m_Field)
                        {
                            // Compare the two folder descriptions.
                            retVal = this.SingleFieldCompare(m_DocFldr1.Description, m_DocFldr2.Description);
                        }
                        else if (Field.Name == m_Field)
                        {
                            // Compare the two folder names.
                            retVal = this.SingleFieldCompare(m_DocFldr1.Name, m_DocFldr2.Name);
                        }
                        else if (Field.Owner == m_Field)
                        {
                            // Compare the two folder owner names.
                            retVal = this.SingleFieldCompare(m_DocFldr1.Owner.Name, m_DocFldr2.Owner.Name);
                        }
                        else if (Field.Client_matter == m_Field) // default
                        {
                            // Query for profile attributes Custom1 & Custom2, and
                            //	compare the parent and child values of both folders.		
                            retVal = this.DoubleFieldCompare(
                                        (this.getAttributeValue(m_DocFldr1.AdditionalProperties, strCust1)),
                                        (this.getAttributeValue(m_DocFldr2.AdditionalProperties, strCust1)),
                                        (this.getAttributeValue(m_DocFldr1.AdditionalProperties, strCust2)),
                                        (this.getAttributeValue(m_DocFldr2.AdditionalProperties, strCust2)));
                        }
                        break;
                    case imObjectType.imTypeDocumentSearchFolder:
                        // It's document folder
                        IManDocumentSearchFolder m_DocSrchFldr1 = (IManDocumentSearchFolder)object1;
                        IManDocumentSearchFolder m_DocSrchFldr2 = (IManDocumentSearchFolder)object2;

                        // Sort by the field user has selected.
                        if (Field.Description == m_Field)
                        {
                            // Compare the two folder descriptions.
                            retVal = this.SingleFieldCompare(m_DocSrchFldr1.Description, m_DocSrchFldr2.Description);
                        }
                        else if (Field.Name == m_Field)
                        {
                            // Compare the two folder names.
                            retVal = this.SingleFieldCompare(m_DocSrchFldr1.Name, m_DocSrchFldr2.Name);
                        }
                        else if (Field.Owner == m_Field)
                        {
                            // Compare the two folder owner names.
                            retVal = this.SingleFieldCompare(m_DocSrchFldr1.Owner.Name, m_DocSrchFldr2.Owner.Name);
                        }
                        else if (Field.Client_matter == m_Field) // default
                        {
                            // Query for profile attributes Custom1 & Custom2, and
                            //	compare the parent and child values of both folders.		
                            retVal = this.DoubleFieldCompare(
                                        (this.getAttributeValue(m_DocSrchFldr1.AdditionalProperties, strCust1)),
                                        (this.getAttributeValue(m_DocSrchFldr2.AdditionalProperties, strCust1)),
                                        (this.getAttributeValue(m_DocSrchFldr1.AdditionalProperties, strCust2)),
                                        (this.getAttributeValue(m_DocSrchFldr2.AdditionalProperties, strCust2)));
                        }
                        break;
                    case imObjectType.imTypeFolderShortcut:
                        // It's folder shortcut
                        IManFolderShortcut m_FS1 = (IManFolderShortcut)object1;
                        IManFolderShortcut m_FS2 = (IManFolderShortcut)object2;

                        // Sort by the field user has selected.
                        if (Field.Description == m_Field)
                        {
                            // Compare the two shortcut descriptons.
                            retVal = this.SingleFieldCompare(m_FS1.Description, m_FS2.Description);
                        }
                        else if (Field.Name == m_Field)
                        {
                            // Compare the two shortcut names.
                            retVal = this.SingleFieldCompare(m_FS1.Name, m_FS2.Name);
                        }
                        else if (Field.Owner == m_Field)
                        {
                            // Compare the two shortcut owner names.
                            retVal = this.SingleFieldCompare(m_FS1.Owner.Name, m_FS2.Owner.Name);
                        }
                        else if (Field.Client_matter == m_Field) // default
                        {
                            // Query for profile attributes Custom1 & Custom2, and
                            //	compare the parent and child values of both folders.		
                            retVal = this.DoubleFieldCompare(
                                        (this.getAttributeValue(m_FS1.AdditionalProperties, strCust1)),
                                        (this.getAttributeValue(m_FS2.AdditionalProperties, strCust1)),
                                        (this.getAttributeValue(m_FS1.AdditionalProperties, strCust2)),
                                        (this.getAttributeValue(m_FS2.AdditionalProperties, strCust2)));
                        }
                        break;
                    case imObjectType.imTypeWorkspace:
                        // It's workspace
                        IManWorkspace m_workspace1 = (IManWorkspace)object1;
                        IManWorkspace m_workspace2 = (IManWorkspace)object2;

                        // Sort by the field user has selected.
                        if (Field.Description == m_Field)
                        {
                            // Compare the two workspace descriptons.
                            retVal = this.SingleFieldCompare(m_workspace1.Description, m_workspace2.Description);
                        }
                        else if (Field.Name == m_Field)
                        {
                            // Compare the two workspace names.
                            retVal = this.SingleFieldCompare(m_workspace1.Name, m_workspace2.Name);
                        }
                        else if (Field.Owner == m_Field)
                        {
                            // Compare the two workspace owners.
                            retVal = this.SingleFieldCompare(m_workspace1.Owner.Name, m_workspace2.Owner.Name);
                        }
                        else if (Field.Client_matter == m_Field) // default
                        {
                            // Compare the parent and child values of both workspaces.		
                            retVal = this.DoubleFieldCompare
                                    (m_workspace1.GetAttributeValueByID(imProfileAttributeID.imProfileCustom1).ToString(),
                                        m_workspace2.GetAttributeValueByID(imProfileAttributeID.imProfileCustom1).ToString(),
                                        m_workspace1.GetAttributeValueByID(imProfileAttributeID.imProfileCustom2).ToString(),
                                        m_workspace2.GetAttributeValueByID(imProfileAttributeID.imProfileCustom2).ToString());
                        }
                        break;
                }
                return retVal;
            }
            catch (Exception e)
            {
              
                string msg = e.Message;
                return false;
            }

            // END EXAMPLE OF IMPLEMENTATION
        }

        #endregion
    }
}
