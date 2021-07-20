module Main (main) where

import Data.Array.IArray ((//))
import qualified Data.HashMap.Strict as HM (fromList)
import Data.Maybe (fromJust)

import Text.Docx (Block (..), Block' (..), DocProps (..), Docx (..),
  Marginal (..), MarginalType (..), NumberFormat (..), ParaProps (..), Run (..),
  Run' (..), RunContent (..), Section (..), SectionProps (..), Spacing (..),
  Style(..), Styles(..), ptToTwip, writeDocx)

import Text.Docx.Types.Defaults (bold, defaultDocProps, defaultParaProps,
  defaultSectionMarginals, defaultSectionProps, endnoteReferenceStyle,
  endnoteTextStyle, footnoteReferenceStyle, footnoteTextStyle, italic,
  qtrCmDefaultTabStop, normalStyle)

main :: IO ()
main = writeDocx "example.docx" example

example :: Docx
example = Docx dProps styles sections headers footers
 where
  dProps = qtrCmDefaultTabStop $ defaultDocProps
            { dPrEnPrNumFmt = Just CardinalText
            -- Although this property is set, it does not appear to be respected
            -- by Word for Microsoft 365 (16.0.14131.20278). So, the
            -- corresponding section property is also set in this example.
            , dPrFnPrNumFmt = Just Decimal
            -- Although this property is set, it does not appear to be respected
            -- by Word for Microsoft 365 (16.0.14131.20278). So, the
            -- corresponding section property is also set in this example.
            }
  styles = Styles
    { stylesPPrDefault = Just defaultParaProps
    , stylesRPrDefault = Nothing
    , stylesStyles = HM.fromList
        [ normalStyle
        , footnoteTextStyle
        , footnoteReferenceStyle
        , endnoteTextStyle
        , endnoteReferenceStyle
        , bodyTextStyle
        , headerStyle
        ]
    }
  bodyTextStyle =
    let (_, normal)   = normalStyle
        normalPprops  = fromJust $ pStyleParaProps normal
        normalSpacing = fromJust $ pPrSpacing normalPprops
    in  ( "Body Text"
        , normal
            { pStyleBasedOn = Just "Normal"
            , pStyleParaProps = Just $ normalPprops
                { pPrSpacing = Just $ normalSpacing
                    { spacingAfter = Just $ ptToTwip 7.0 }
                }
            }
        )
  headerStyle =
    let (_, normal) = normalStyle
    in  ( "Header"
        , normal
            { pStyleBasedOn = Just "Normal"}
        )
  sections = [section1, section2]
  section1 = Section s1props
    [ Paragraph (Just "Body Text") (ParaProps Nothing Nothing Nothing [])
        [ Run Nothing mempty
            [ RunText $ "This is an example .docx document. The footnotes " <>
                      "are numbered using 'cardinal text', which is an " <>
                      "option that cannot be selected in Microsoft Word." ]
        , Footnote (Just "Footnote Reference") mempty
            [ Paragraph' (Just "Footnote Text") (ParaProps Nothing Nothing Nothing [])
                [ Run' Nothing mempty
                    [ RunText " This is a footnote."]
                ]
            ]
        ]
   , Paragraph (Just "Body Text") (ParaProps Nothing Nothing Nothing [])
        [ Run Nothing bold
            [ RunText "This is an example .docx document (bold)." ]
        , Footnote (Just "Footnote Reference") mempty
            [ Paragraph' (Just "Footnote Text") (ParaProps Nothing Nothing Nothing [])
                [ Run' Nothing mempty
                    [ RunText " This is a second footnote." ]
                ]
            ]
        ]
   , Paragraph (Just "Body Text") (ParaProps Nothing Nothing Nothing [])
        [ Run Nothing italic
            [ RunText "This is an example .docx document (italic)." ]
        ]
   , Paragraph (Just "Body Text") (ParaProps Nothing Nothing Nothing [])
        [ Run Nothing (bold <> italic)
            [ RunText "This is an example .docx document (bold and italic)." ]
        ]
    ]
  s1props = addNoteNumFmt $ defaultSectionProps
    { sPrMarginals = defaultSectionMarginals //
        [ ((Header, DefaultMarginal), Just 1) ]
    }
  section2 = Section s2props
    [ Paragraph (Just "Body Text") (ParaProps Nothing Nothing Nothing [])
        [ Run Nothing mempty
            [ RunText $ "This is a second section of the document. The " <>
                      "endnotes are numbered using 'decimal'." ]
        , Endnote (Just "Endnote Reference") mempty
            [ Paragraph' (Just "Endnote Text") (ParaProps Nothing Nothing Nothing [])
                [ Run' Nothing mempty
                    [ RunText " This is an endnote." ]
                ]
            ]
        ]

    ]
  s2props = addNoteNumFmt defaultSectionProps
  addNoteNumFmt props = props
    { sPrFnPrNumFmt = Just CardinalText
    , sPrEnPrNumFmt = Just Decimal
    }
  headers = HM.fromList
    [ ( 1
      , [ Paragraph' (Just "Header") (ParaProps Nothing Nothing Nothing [])
            [ Run' Nothing mempty
                [ RunText "example.docx example" ]
            ]
        ]
      )
    ]
  footers = mempty
