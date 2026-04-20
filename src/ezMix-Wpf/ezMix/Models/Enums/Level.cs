using System.ComponentModel;

namespace ezMix.Models.Enums
{
    public enum Level
    {
        [Description("NB")]
        Know,
        [Description("TH")]
        Understand,
        [Description("VD")]
        Manipulate,
        [Description("")]
        None
    }
}
