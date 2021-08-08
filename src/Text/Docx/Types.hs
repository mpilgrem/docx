{-# LANGUAGE DeriveGeneric #-}

{-|
Module      : Text.Docx.Types
Description : Docx types and defaults
Copyright   : (c) Mike Pilgrem, 2021
License     : BSD-3
Maintainer  : public@pilgrem.com
Stability   : experimental
Portability : POSIX, Windows

The type representing an Office Open XML document in the Wordprocessing category
('Docx') and related types.
-}
module Text.Docx.Types
  ( Docx (..)
  , DocProps (..)
  , NumberFormat (..)
  , Twip
  , Emu
  , NumDefs (..)
  , AbstractNum (..)
  , AbstractNumProps (..)
  , NumLevel (..)
  , LevelSuffix (..)
  , Numbering (..)
  , NumberingProps (..)
  , Styles (..)
  , StyleName
  , Style (..)
  , ParaProps (..)
  , NumProps (..)
  , Justification (..)
  , Spacing (..)
  , LineRule (..)
  , Indentation (..)
  , Tab
  , TabStyle (..)
  , RunProps (..)
  , HalfPt
  , Toggle
  , Fonts (..)
  , Section (..)
  , SectionProps (..)
  , PageSize
  , PageMargins (..)
  , Block (..)
  , Block' (..)
  , Run (..)
  , Run' (..)
  , RGB (..)
  , Underline (..)
  , UnderlinePattern (..)
  , VertAlign (..)
  , RunContent (..)
  , BreakType (..)
  , ClearType (..)
  , TableProps (..)
  , TableLayout (..)
  , CellMargins (..)
  , CellMargin (..)
  , TableGrid
  , GridCol
  , Row (..)
  , Row' (..)
  , RowProps (..)
  , Cell (..)
  , Cell' (..)
  , CellProps (..)
  , CellBorders (..)
  , CellBorder (..)
  , EighthPt
  , Marginal (..)
  , MarginalType (..)
  , NoteType (..)
  ) where

import Data.Array.IArray (Array)
import Data.Ix (Ix)
import Data.Word (Word8)

import Codec.Picture (DynamicImage)
import Data.Colour.RGBSpace (RGB (..))
import Data.HashMap.Strict (HashMap)
import Generics.Deriving (Generic)
import Generics.Deriving.Monoid (First, Last, memptydefault)
import Generics.Deriving.Semigroup (gsappenddefault)

-- |Representation of Office Open XML Wordprocessing documents.
data Docx = Docx
  { docxProps    :: DocProps
    -- ^ Document-wide properties.
  , docxNumDefs :: NumDefs
    -- ^ Automatic numbering definitions that can be referenced in the
    -- document's sections.
  , docxStyles   :: Styles
    -- ^ Styles that can be referenced in the document's sections.
  , docxSections :: [Section]
    -- ^ The series of sections of the document.
  } deriving (Eq)

-- |Representation of document properties.
data DocProps = DocProps
  { dPrFnPrNumFmt :: Maybe NumberFormat
    -- ^ Number format for footnotes. Word for Microsoft 365 (16.0.14131.20278)
    -- appears not to respect this property. See the corresponding
    -- 'SectionProps' field.
  , dPrEnPrNumFmt :: Maybe NumberFormat
    -- ^ Number format for endnotes. Word for Microsoft 365 (16.0.14131.20278)
    -- appears not to respect this property. See the corresponding
    -- 'SectionProps' field.
  , dPrDefaultTabStop :: Maybe Twip
    -- ^ Spacing of default tab stop locations.
  } deriving Eq

-- |Representing of automatic numbering formats.
data NumberFormat
  = Aiueo
  | AiueoFullWidth
  | ArabicAbjad
  | ArabicAlpha
  | BahtText
  | Bullet
  | CardinalText
  | Chicago
  | ChineseCounting
  | ChineseCountingThousand
  | ChineseLegalSimplified
  | Chosung
  | Custom
  | Decimal
  | DecimalEnclosedCircle
  | DecimalEnclosedCircleChinese
  | DecimalEnclosedFullstop
  | DecimalEnclosedParen
  | DecimalFullWidth
  | DecimalFullWidth2
  | DecimalHalfWidth
  | DecimalZero
  | DollarText
  | Ganada
  | Hebrew1
  | Hebrew2
  | Hex
  | HindiConsonants
  | HindiCounting
  | HindiNumbers
  | HindiVowels
  | IdeographDigital
  | IdeographEnclosedCircle
  | IdeographLegalTraditional
  | IdeographTraditional
  | IdeographZodiac
  | IdeographZodiacTraditional
  | Iroha
  | IrohaFullWidth
  | JapaneseCounting
  | JapaneseDigitalTenThousand
  | JapaneseLegal
  | KoreanCounting
  | KoreanDigital
  | KoreanDigital2
  | KoreanLegal
  | LowerLetter
  | LowerRoman
  | NoNumberFormat
  | NumberInDash
  | Ordinal
  | OrdinalText
  | RussianLower
  | RussianUpper
  | TaiwaneseCounting
  | TaiwaneseCountingThousand
  | TaiwaneseDigital
  | ThaiCounting
  | ThaiLetters
  | ThaiNumbers
  | UpperLetter
  | UpperRoman
  | VietnameseCounting
  deriving Eq

-- |Synonym representing measurments in twip (20th of a point, 1,440th of an
-- inch).
type Twip = Int

-- |Synonym representing measurements in EMUs (English Metric Units) (360,000th
-- of a centimetre, 914,400th of an inch).
type Emu = Int

-- |Representation of automatic numbering definitions.
data NumDefs = NumDefs [AbstractNum] [Numbering]
             deriving Eq

data AbstractNum = AbstractNum AbstractNumProps [NumLevel]
                 deriving Eq

data AbstractNumProps = AbstractNumProps
  { anPrAbstractNumId :: Int
  , anPrName          :: String
  } deriving Eq

-- |Representation of numbering levels.
data NumLevel = NumLevel
  { nlIlvl           :: Int  -- ^ 0-based index to numbering levels.
  , nlStart          :: Maybe Int
  , nlNumFmt         :: Maybe NumberFormat
  , nlLvlRestart     :: Maybe Int
  , nlPStyle         :: Maybe String
  , nlIsLgl          :: Maybe Bool
  , nlSuff           :: Maybe LevelSuffix
  , nlLvlText        :: Maybe String
  , nlLvlJc          :: Maybe Justification
  , nlPPr            :: Maybe ParaProps
  , nlRPr            :: Maybe RunProps
  } deriving Eq

data LevelSuffix
  = TabSuffix
  | SpaceSuffix
  | NoSuffix
  deriving Eq

data Numbering = Numbering NumberingProps [NumLevel]
               deriving Eq

data NumberingProps = NumberingProps
  { nPrNumId         :: Int
  , nPrAbstractNumId :: Int
  } deriving Eq

-- |Representation of Style Definitions parts of an Office Open XML document
-- in the Wordprocessing category.
data Styles = Styles
  { stylesPPrDefault :: Maybe ParaProps  -- ^ Default paragraph properties.
  , stylesRPrDefault :: Maybe RunProps   -- ^ Default run properies.
  , stylesStyles     :: HashMap StyleName Style
  } deriving Eq

-- |Synonym representing style names. Styles are identified by their names,
-- which must be unique in a particular Office Open XML Wordprocessing document.
type StyleName = String

-- |Representation of style definitions.
data Style
  = ParaStyle  -- ^ A paragraph style.
      { pStyleBasedOn   :: Maybe StyleName
      , pStyleParaProps :: Maybe ParaProps
      , pStyleRunProps  :: Maybe RunProps
      , pStyleIsDefault :: Bool
      , pStyleIsQFormat :: Bool
      }
  | RunStyle  -- ^ A run (character) style.
      { rStyleBasedOn   :: Maybe StyleName
      , rStyleRunProps  :: Maybe RunProps
      , rStyleIsDefault :: Bool
      , rStyleIsQFormat :: Bool
      }
  deriving Eq

-- |Representation of sections.
data Section = Section SectionProps [Block]
             deriving Eq

-- |Representation of the properties of a section.
data SectionProps = SectionProps
  { sPrPgSz       :: PageSize
  , sPrPgMar      :: PageMargins
  , sPrMarginals  :: Array (Marginal, MarginalType) [Block']
  , sPrFnPrNumFmt :: Maybe NumberFormat
  , sPrEnPrNumFmt :: Maybe NumberFormat
  } deriving Eq

-- |Synonym representing page sizes, (width, height).
type PageSize = (Twip, Twip)

-- |Representation of page margins.
data PageMargins = PageMargins
  { pmTopOverlap    :: Bool
  , pmBottomOverlap :: Bool
  , pmTop           :: Twip
  , pmBottom        :: Twip
  , pmLeft          :: Twip
  , pmRight         :: Twip
  , pmHeader        :: Twip
  , pmFooter        :: Twip
  , pmGutter        :: Twip
  } deriving Eq

-- |Representation of running marginals.
data Marginal
  = Header
  | Footer
  deriving (Bounded, Enum, Eq, Ix, Ord, Show)

-- |Representation of marginal types.
data MarginalType
  = DefaultMarginal
  | FirstMarginal
  | EvenMarginal
  deriving (Bounded, Enum, Eq, Ix, Ord, Show)

-- |Representation of block-level markup in the main document.
--
-- Reference:
-- /ECMA-376-1:2016, Fundamentals and Markup Language Reference §17.2.2/.
data Block
  = Paragraph (Maybe String) ParaProps [Run]
    -- ^ A paragraph.
  | Table TableProps TableGrid [Row]
  deriving Eq

-- |Representation of block-level markup other than in the main document
-- (headers, footers, footnotes and endnotes).
--
-- Reference:
-- /ECMA-376-1:2016, Fundamentals and Markup Language Reference §17.2.2/.
data Block'
  = Paragraph' (Maybe String) ParaProps [Run']
    -- ^ A paragraph.
  | Table' TableProps TableGrid [Row']
  deriving Eq

-- |Representation of a paragraph's properties.
data ParaProps = ParaProps
  { pPrNumPr   :: Maybe NumProps
  , pPrTabs    :: [Tab]
  , pPrSpacing :: Maybe Spacing
  , pPrInd     :: Maybe Indentation
  , pPrJc      :: Maybe Justification
  } deriving Eq

data NumProps = NumProps
  { numPrIlvl  :: Maybe Int
  , numPrNumId :: Maybe Int
  } deriving Eq

-- |Representation of paragraph spacings.
data Spacing = Spacing
  { spacingBefore :: Maybe Twip
  , spacingAfter  :: Maybe Twip
  , spacingLine   :: Maybe (LineRule, Int)
  } deriving Eq

-- |Representation of a paragraph line rules.
data LineRule
  = Auto
  | Exactly
  | AtLeast
  deriving Eq

-- |Representation of a paragraph's justifications and alignments.
data Justification
  = Start
  | End
  | Center
  | Both
  | Distribute
  deriving Eq

-- |Representation of a paragraph's indentations.
data Indentation = Indentation
  { indStart     :: Maybe Twip
  , indEnd       :: Maybe Twip
  , indHanging   :: Maybe Twip
  , indFirstLine :: Maybe Twip
  } deriving Eq

-- |Synonym representing tab stops.
type Tab = (TabStyle, Twip)

-- |Representation of tab styles.
data TabStyle
  = BarTab
  | CenterTab
  | ClearTab
  | DecimalTab
  | EndTab
  | StartTab
  deriving Eq

-- |Representation of a table's properties.
data TableProps = TableProps
  { tblPrTblW      :: Maybe Int
  , tblPrTblLayout :: Maybe TableLayout
  , tblPrCellMar   :: Maybe CellMargins
  } deriving Eq

-- |Representation of a table's layout types.
data TableLayout
  = AutoLayout
  | FixedLayout
  deriving Eq

-- |Representation of a table row cell's margins.
data CellMargins = CellMargins
  { cmTop    :: Maybe CellMargin
  , cmBottom :: Maybe CellMargin
  , cmStart  :: Maybe CellMargin
  , cmEnd    :: Maybe CellMargin
  } deriving Eq

-- |Representation of cell margins.
newtype CellMargin = CellMargin
  { tcMarginW :: Int
  } deriving Eq

-- |Synonym representing a table's grids.
type TableGrid = [GridCol]

type GridCol = Int

-- |Representation of runs of text in the main document.
data Run
  = Run (Maybe StyleName) RunProps [RunContent]
  | Footnote (Maybe StyleName) RunProps [Block']
  | Endnote (Maybe StyleName) RunProps [Block']
  deriving Eq

-- |Representation of runs of text other than in the main document.
data Run' = Run' (Maybe StyleName) RunProps [RunContent]
         deriving Eq

-- |Representation of a run's properties.
data RunProps = RunProps
  { rPrIsBold      :: Toggle
  , rPrIsItalic    :: Toggle
  , rPrU           :: Last Underline
  , rPrIsStrike    :: Toggle
  , rPrIsDStrike   :: Toggle
  , rPrIsCaps      :: Toggle
  , rPrIsSmallCaps :: Toggle
  , rPrIsEmboss    :: Toggle
  , rPrIsImprint   :: Toggle
  , rPrIsOutline   :: Toggle
  , rPrIsShadow    :: Toggle
  , rPrColor       :: Last (RGB Word8)
  , rPrRFonts      :: Last Fonts
  , rPrSz          :: Last HalfPt
  , rPrVertAlign   :: Last VertAlign
  } deriving (Eq, Generic, Show)

-- |Synonym representing halves of a point
type HalfPt = Int

instance Semigroup RunProps where

  (<>) = gsappenddefault

instance Monoid RunProps where

  mempty = memptydefault

-- |Synonyn representing toggle properties
type Toggle = First Bool

-- |Representation of a run's fonts.
data Fonts = Fonts
  { rFontsAscii    :: Maybe String
  , rFontsCs       :: Maybe String
  , rFontsHAnsi    :: Maybe String
  , rFontsEastAsia :: Maybe String
  } deriving (Eq, Show)

-- |Representation of styles of underline.
data Underline = Underline (Maybe (RGB Word8)) UnderlinePattern
               deriving (Eq, Show)

-- |Representation of patterns of underline.
data UnderlinePattern
  = Dash
  | DashDotDotHeavy
  | DashDotHeavy
  | DashedHeavy
  | DashLong
  | DashLongHeavy
  | DotDash
  | DotDotDash
  | Dotted
  | DottedHeavy
  | Double
  | NoUnderline
  | Single
  | Thick
  | Wave
  | WavyDouble
  | WavyHeavy
  | Words
  deriving (Eq, Show)

-- |Representation of run vertical alignments.
data VertAlign
  = Baseline
  | Superscript
  | Subscript
  deriving (Eq, Show)

-- |Representation of run contents in the main document.
data RunContent
  = RunText String
  | Break BreakType
  | NoBreakHyphen
  | SoftHyphen
  | Symbol String Int
  | Tab
  | Drawing DynamicImage  -- ^ Ignored outside of the main document.
  deriving Eq

-- |Representation of break types
data BreakType
  = TextWrapping ClearType
  | ColumnBreak
  | PageBreak
  deriving (Eq, Show)

-- |Representation of clear types
data ClearType
  = NoneClear
  | LeftClear
  | RightClear
  | AllClear
  deriving (Eq, Show)

-- |Representation of table rows in the main document.
data Row = Row RowProps [Cell] deriving Eq

-- |Representation of table rows other than in the main document.
data Row' = Row' RowProps [Cell'] deriving Eq

-- |Representation of table row properties.
data RowProps = RowProps
  { trPrCantSplit :: Maybe Bool
  , trPrTblHeader :: Maybe Bool
  } deriving (Eq, Show)

-- |Representation of table row cells in the main document.
data Cell = Cell CellProps [Block] deriving Eq

-- |Representation of table row cells other than in the main document.
data Cell' = Cell' CellProps [Block'] deriving Eq

-- |Representation of table cells properties.
data CellProps = CellProps
  { tcPrTcW       :: Twip
  , tcPrTcBorders :: Maybe CellBorders
  } deriving (Eq, Show)

-- |Representation of table cells borders.
data CellBorders = CellBorders
  { cbTop    :: Maybe CellBorder
  , cbBottom :: Maybe CellBorder
  , cbLeft   :: Maybe CellBorder
  , cbRight  :: Maybe CellBorder
  } deriving (Eq, Show)

-- |Represention of cell borders.
newtype CellBorder = CellBorder
  { tcBorderSz :: EighthPt
  } deriving (Eq, Show)

-- |Synonym representing eighths of a point.
type EighthPt = Int

-- |Representation of footnote or endnote types.
data NoteType
  = NormalNote
  | Separator
  | ContinuationSeparator
  deriving (Eq, Show)
