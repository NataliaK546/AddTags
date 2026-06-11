using System.Windows;

namespace WpfAddTags
{
    public class TagInfo
    {
        public string Item {  get; set; }
        public string Name { get; set; }
        public string Type;
        public string TypeRW;
        public string AWU;
        public string Legend;
        public string NumStation;
        public int NumTag;
        public string CAAG;
        public string Koef;
        public string Color;
        public string Init;
        public string Group;
        public string ChangeVal;
        public string SaveByTime;
        public string RarelyChanging;

        public TagInfo(object[,] worksheet, int j, int i)
        {
            Item = worksheet[j, i + 1] == null ? "" : worksheet[j, i + 1].ToString();
            Name = worksheet[j, i + 2] == null ? "" : worksheet[j, i + 2].ToString();
            Type = worksheet[j, i + 3] == null ? "" : worksheet[j, i + 3].ToString();
            TypeRW = worksheet[j, i + 4] == null ? "" : worksheet[j, i + 4].ToString();
            AWU = worksheet[j, i + 5] == null ? "" : worksheet[j, i + 5].ToString();
            Legend = worksheet[j, i + 6] == null ? "" : worksheet[j, i + 6].ToString();
            NumStation = worksheet[j, i + 7] == null ? "" : worksheet[j, i + 7].ToString();

            var stringnumtag = worksheet[j, i + 8] == null ? string.Empty: worksheet[j, i + 8].ToString();
            var resintnumtag = int.TryParse(stringnumtag, out int intnumtag);
            if (resintnumtag) NumTag = intnumtag;
            else
            {
                NumTag = 0;
                MessageBox.Show("Неправильная конвертация numTag; " + worksheet[j, i + 8]);
            }

            CAAG = worksheet[j, i + 9] == null ? "" : worksheet[j, i + 9].ToString();
            Koef = worksheet[j, i + 10] == null ? "" : worksheet[j, i + 10].ToString();
            Color = worksheet[j, i + 11] == null ? "" : worksheet[j, i + 11].ToString();
            Init = worksheet[j, i + 12] == null ? "" : worksheet[j, i + 12].ToString();
            Group = worksheet[j, i + 13] == null ? "" : worksheet[j, i + 13].ToString();
            ChangeVal = worksheet[j, i + 14] == null ? "" : worksheet[j, i + 14].ToString();
            SaveByTime = worksheet[j, i + 15] == null ? "" : worksheet[j, i + 15].ToString();
            RarelyChanging = worksheet[j, i + 16] == null ? "" : worksheet[j, i + 16].ToString();
        }
        public override bool Equals(object obj)
        {
            if (obj is TagInfo tag)
                return Item == tag.Item;
            return false;
        }
        public override int GetHashCode() => Item.GetHashCode();
    }
}
