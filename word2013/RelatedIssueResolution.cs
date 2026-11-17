using System;

namespace RadiologyAgent {
  // Pure, conservative test for whether a comment was actually fixed by an
  // exact replacement. Unrelated comments must remain open.
  internal static class RelatedIssueResolution {
    internal static bool IsResolved(int noteStart,int noteEnd,string noteText,
        int editStart,int editEnd,string replacement,bool isChosenIssue) {
      if(isChosenIssue)return true;
      if(noteStart<editStart||noteEnd>editEnd||noteStart>=noteEnd)
        return false;
      if(String.IsNullOrWhiteSpace(noteText)||replacement==null)return false;
      // A whole-sentence replacement can correct multiple overlapping notes,
      // but only if that note's exact old anchor is absent afterwards.
      return replacement.IndexOf(noteText,StringComparison.Ordinal)<0;
    }
    internal static bool SelfTest() {
      string edit="Elevated left diaphragmatic dome with lower lobar bands.";
      return IsResolved(14,34,"diaphragmatic copula",0,90,edit,false)
        && IsResolved(14,34,"diaphragmatic copula",14,34,"diaphragmatic dome",true)
        && !IsResolved(4,15,"left",0,90,edit,false)
        && !IsResolved(91,105,"unrelated",0,90,edit,false)
        && IsResolved(4,20,"renal",0,90,"Renal",false);
    }
  }
}
