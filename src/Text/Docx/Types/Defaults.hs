{-|
Module      : Text.Docx.Types.Defaults
Description : Docx types and defaults
Copyright   : (c) Mike Pilgrem, 2021
License     : BSD-3
Maintainer  : public@pilgrem.com
Stability   : experimental
Portability : POSIX, Windows

Default values for types representing Office Open XML Wordpressing documents.
-}
module Text.Docx.Types.Defaults
  ( defaultDocProps
  , qtrCmDefaultTabStop
  , defaultNumDefs
  , normalStyle
  , footnoteTextStyle
  , endnoteTextStyle
  , captionStyle
  , footnoteReferenceStyle
  , endnoteReferenceStyle
  , defaultSectionProps
  , a4PageSize
  , defaultPageMargins
  , defaultSectionMarginals
  , defaultParaProps
  , bold
  , italic
  , calibri
  , defaultTableProps
  , zeroCellMargins
  , zeroCellMargin
  , defaultRowProps
  ) where

import Data.Array.IArray (Array, array)
import Data.Monoid (Last (..))

import Text.Docx.Types
import Text.Docx.Utilities

defaultDocProps :: DocProps
defaultDocProps       = DocProps
  { dPrFnPrNumFmt     = Nothing
  , dPrEnPrNumFmt     = Nothing
  , dPrDefaultTabStop = Nothing
  }

qtrCmDefaultTabStop :: DocProps -> DocProps
qtrCmDefaultTabStop props = props
  { dPrDefaultTabStop = Just $ cmToTwip 0.25 }

defaultNumDefs :: NumDefs
defaultNumDefs = NumDefs [] []

normalStyle :: (StyleName, Style)
normalStyle = ("Normal", ParaStyle
  { pStyleBasedOn   = Nothing
  , pStyleParaProps = Just defaultParaProps
  , pStyleRunProps  = Nothing
  , pStyleIsDefault = True
  , pStyleIsQFormat = True
  })

footnoteTextStyle :: (StyleName, Style)
footnoteTextStyle = ("Footnote Text", (snd normalStyle)
  { pStyleBasedOn   = Just (fst normalStyle)
  })

endnoteTextStyle :: (StyleName, Style)
endnoteTextStyle = ("Endnote Text", (snd normalStyle)
  { pStyleBasedOn   = Just (fst normalStyle)
  })

captionStyle :: (StyleName, Style)
captionStyle = ("Caption", (snd normalStyle)
  { pStyleBasedOn = Just (fst normalStyle)
  , pStyleRunProps = Just $ mempty { rPrIsBold = toggle True }
  , pStyleIsDefault = False
  })

footnoteReferenceStyle :: (StyleName, Style)
footnoteReferenceStyle = ("Footnote Reference", RunStyle
  { rStyleBasedOn   = Nothing
  , rStyleRunProps  = Just $ mempty { rPrVertAlign = Last $ Just Superscript }
  , rStyleIsDefault = True
  , rStyleIsQFormat = True
  })

endnoteReferenceStyle :: (StyleName, Style)
endnoteReferenceStyle = ("Endnote Reference", RunStyle
  { rStyleBasedOn   = Nothing
  , rStyleRunProps  = Just $ mempty { rPrVertAlign = Last $ Just Superscript }
  , rStyleIsDefault = True
  , rStyleIsQFormat = True
  })

defaultSectionProps :: SectionProps
defaultSectionProps = SectionProps
  { sPrPgSz             = a4PageSize
  , sPrPgMar            = defaultPageMargins
  , sPrMarginals        = defaultSectionMarginals
  , sPrFnPrNumFmt       = Nothing
  , sPrEnPrNumFmt       = Nothing
  }

defaultSectionMarginals :: Array (Marginal, MarginalType) [Block']
defaultSectionMarginals =
  array ((minBound, minBound), (maxBound, maxBound))
        [((m, mt), []) | m <- [minBound .. maxBound]
                            , mt <- [minBound .. maxBound]
                            ]

a4PageSize :: PageSize
a4PageSize = (cmToTwip 21.0, cmToTwip 29.7)

defaultPageMargins :: PageMargins
defaultPageMargins = PageMargins
  { pmTopOverlap    = False
  , pmBottomOverlap = False
  , pmTop           = cmToTwip 4.50
  , pmBottom        = cmToTwip 4.50
  , pmLeft          = cmToTwip 5.00
  , pmRight         = cmToTwip 3.00
  , pmHeader        = cmToTwip 1.50
  , pmFooter        = cmToTwip 1.50
  , pmGutter        = cmToTwip 0.00
  }

defaultParaProps :: ParaProps
defaultParaProps = ParaProps
  { pPrNumPr = Nothing
  , pPrTabs = []
  , pPrSpacing = Just defaultSpacing
  , pPrInd     = Nothing
  , pPrJc      = Nothing

  }

defaultSpacing :: Spacing
defaultSpacing = Spacing
  { spacingBefore = Just 0
  , spacingAfter  = Just 0
  , spacingLine   = Just (AtLeast, ptToTwip 14.0)
  }

defaultTableProps :: TableProps
defaultTableProps = TableProps
  { tblPrTblW      = Nothing
  , tblPrTblLayout = Nothing
  , tblPrCellMar   = Nothing
  }

zeroCellMargins :: CellMargins
zeroCellMargins = CellMargins
  { cmTop    = Just zeroCellMargin
  , cmBottom = Just zeroCellMargin
  , cmStart  = Just zeroCellMargin
  , cmEnd    = Just zeroCellMargin
  }

zeroCellMargin :: CellMargin
zeroCellMargin = CellMargin { tcMarginW = 0 }

italic :: RunProps
italic = mempty {rPrIsItalic = toggle True}

bold :: RunProps
bold = mempty {rPrIsBold = toggle True}

calibri :: Fonts
calibri = Fonts
  { rFontsAscii    = Just "Calibri"
  , rFontsCs       = Nothing
  , rFontsHAnsi    = Nothing
  , rFontsEastAsia = Nothing
  }

defaultRowProps :: RowProps
defaultRowProps = RowProps
  { trPrCantSplit = Nothing
  , trPrTblHeader = Nothing
  }
