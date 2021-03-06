{-# LANGUAGE FlexibleInstances #-}
{-|
Module      : Text.Docx.Writer
Description : Write .docx file
Copyright   : (c) Mike Pilgrem, 2021
License     : BSD-3
Maintainer  : public@pilgrem.com
Stability   : experimental
Portability : POSIX, Windows

-}
module Text.Docx.Writer
  ( docx
  , writeDocx
  ) where

import Data.Array.IArray (Array, assocs)
import Data.List (foldl', mapAccumL)
import Data.Maybe (fromMaybe, maybeToList)
import Data.Monoid (First (..), Last (..))
import Data.Word (Word8)
import Numeric (showHex)

import Codec.Archive.Zip (Archive, Entry, addEntryToArchive, emptyArchive,
  fromArchive, toEntry)
import Codec.Picture (DynamicImage (..), Image (..), dynamicMap)
import Codec.Picture.Saving (imageToPng)
import Control.Monad.Extra (concatMapM)
import Control.Monad.State (State, evalState, gets, modify, runState)
import qualified Data.ByteString.Lazy.Char8 as BS
import Data.HashMap.Lazy (toList)
import Data.Time.Clock.System (SystemTime (systemSeconds), getSystemTime)
import Text.XML.Light
    ( unqual,
      parseXMLDoc,
      ppTopElement,
      blank_name,
      Node(node),
      Attr(Attr),
      Content(Elem),
      Element,
      QName(QName, qName, qPrefix) )

import Text.Docx.Types
import Text.Docx.Types.Defaults (defaultParaProps)
import Text.Docx.Utilities (pxToEmu)

-- |Create a file from an Office Open XML Wordpressing document.
writeDocx :: FilePath -> Docx -> IO ()
writeDocx fp d = do
  sysTime <- getSystemTime
  BS.writeFile fp (fromArchive $ docx d sysTime)

--------------------------------------------------------------------------------
-- Assemble the archive from parts
--------------------------------------------------------------------------------

data OtherParts = OtherParts
  { opFootnoteNo     :: Int
  , opFootnotes      :: [(Int, [Block'])]
  , opEndnoteNo      :: Int
  , opEndnotes       :: [(Int, [Block'])]
  , opHeaderFooterNo :: Int
  , opHeaders        :: [(Int, [Block'])]
  , opFooters        :: [(Int, [Block'])]
  , opDrawingNo      :: Int
  , opDrawings       :: [(Int, DynamicImage)]
  } deriving Eq

emptyOtherParts :: OtherParts
emptyOtherParts = OtherParts
  { opFootnoteNo     = 1  -- Items 0 and 1 used for separators.
  , opFootnotes      = []
  , opEndnoteNo      = 1  -- Items 0 and 1 used for separators.
  , opEndnotes       = []
  , opHeaderFooterNo = 5  -- Items 1 to 5 used for styles, settings, footnotes,
                          -- endnotes and numbering.
  , opHeaders        = []
  , opFooters        = []
  , opDrawingNo      = 0
  , opDrawings       = []
  }

-- |Create an archive from an Office Open XML Wordprocessing document, given a
-- system time.
docx :: Docx -> SystemTime -> Archive
docx d sysTime = foldl' (flip addEntryToArchive) emptyArchive entries
 where
  sysTime' = toInteger $ systemSeconds sysTime
  (documentEntry', otherParts) = documentEntry d sysTime'
  hasNumbering = case d of Docx _ (NumDefs [] _) _ _ -> False; _ -> True
  footnotes = opFootnotes otherParts
  hasFootnotes = not $ null footnotes
  endnotes = opEndnotes otherParts
  hasEndnotes = not $ null endnotes
  headers = opHeaders otherParts
  hRefCount = length headers
  footers = opFooters otherParts
  fRefCount = length footers
  drawings = opDrawings otherParts
  entries = map (\f -> f sysTime') (
    [ contentTypesEntry hasNumbering hRefCount fRefCount
                        hasFootnotes hasEndnotes

    , relsEntry
    , documentXmlRelsEntry hasNumbering headers footers hasFootnotes hasEndnotes
                           drawings
    , stylesEntry d
    , settingsEntry (docxProps d) hasFootnotes hasEndnotes
    ] <>
    [ numberingEntry d | hasNumbering ] <>
    imageEntries (opDrawings otherParts) <>
    concatMap (uncurry marginalEntries) [ (Header, headers), (Footer, footers) ]
    ) <>
    maybeToList (footnotesEntry (opFootnotes otherParts) sysTime') <>
    maybeToList (endnotesEntry (opEndnotes otherParts) sysTime') <>
    [documentEntry']

--------------------------------------------------------------------------------
-- Archive entries
--------------------------------------------------------------------------------

-- |Create the [Content_Types].xml entry
contentTypesEntry :: Bool -> Int -> Int -> Bool -> Bool -> Integer -> Entry
contentTypesEntry hasNumbering hRefCount fRefCount hasFootnotes hasEndnotes sysTime =
  toEntry "[Content_Types].xml"
          sysTime
          (BS.pack $ ppTopElement $
             contentTypes hasNumbering hRefCount fRefCount
                                       hasFootnotes hasEndnotes)

-- |Create the _rels/.rels entry
relsEntry :: Integer -> Entry
relsEntry sysTime = toEntry
  "_rels/.rels"
  sysTime
  (BS.pack $ ppTopElement rels)

-- |Create the _rels/document.xml.rels entry
documentXmlRelsEntry :: Bool
                     -> [(Int, [Block'])]
                     -> [(Int, [Block'])]
                     -> Bool
                     -> Bool
                     -> [(Int, DynamicImage)]
                     -> Integer -> Entry
documentXmlRelsEntry
  hasNumbering headers footers hasFootnotes hasEndnotes drawings sysTime =
    toEntry
      "_rels/document.xml.rels"
      sysTime
      (BS.pack $ ppTopElement $
         documentXmlRels hasNumbering headers footers hasFootnotes hasEndnotes drawings)

-- |Create the settings.xml entry
settingsEntry :: DocProps
              -> Bool  -- ^ Has footnotes?
              -> Bool  -- ^ Has endnotes?
              -> Integer -> Entry
settingsEntry props hasFootnotes hasEndnotes sysTime = toEntry
  "settings.xml"
  sysTime
  (BS.pack $ ppTopElement $ settingsXml props hasFootnotes hasEndnotes)

-- |Create the numbering.xml entry
numberingEntry :: Docx -> Integer -> Entry
numberingEntry (Docx _ numDefs _ _) sysTime = toEntry
  "numbering.xml"
  sysTime
  (BS.pack $ ppTopElement $ numberingXml numDefs)

-- |Create the styles.xml entry
stylesEntry :: Docx -> Integer -> Entry
stylesEntry d sysTime = toEntry
  "styles.xml"
  sysTime
  (BS.pack $ ppTopElement $ stylesXml $ docxStyles d)

imageEntries :: [(Int, DynamicImage)] -> [Integer -> Entry]
imageEntries images = snd $ mapAccumL imageEntry 1 images
 where
  imageEntry :: Int -> (a, DynamicImage) -> (Int, Integer -> Entry)
  imageEntry n (_, di) = (n + 1, \st -> toEntry
    ("media/image" <> show n <> ".png")
    st
    (imageToPng di))

-- |Create the document.xml entry. NB Microsoft Word for Microsoft 365 will put
-- this in folder \word.
documentEntry :: Docx -> Integer -> (Entry, OtherParts)
documentEntry d sysTime =
  let (documentElement, otherParts) = documentXml d
      documentByteString = BS.pack $ ppTopElement documentElement
  in  (toEntry "document.xml" sysTime documentByteString, otherParts)

-- |Representation of types of note
data Note
  = FootNote
  | EndNote
  deriving (Eq, Show)

noteType :: Note -> String
noteType nt = case nt of
  FootNote -> "footnote"
  EndNote  -> "endnote"

noteTypes :: Note -> String
noteTypes nt = noteType nt <> "s"

noteTypeXml :: Note -> FilePath
noteTypeXml nt = noteTypes nt <> ".xml"

-- |Create the footnotes.xml entry.
footnotesEntry :: [(Int, [Block'])] -> Integer -> Maybe Entry
footnotesEntry = notesEntry FootNote

-- |Create the endnotes.xml entry.
endnotesEntry :: [(Int, [Block'])] -> Integer -> Maybe Entry
endnotesEntry = notesEntry EndNote

-- |Create the footnotes.xml or endnotes.xml entry.
notesEntry :: Note -> [(Int, [Block'])] -> Integer -> Maybe Entry
notesEntry _ [] _ = Nothing
notesEntry nt ns sysTime =
  Just $ toEntry (noteTypeXml nt) sysTime $ BS.pack $ ppTopElement notesElement
 where
  notesWName = wName $ noteTypes nt
  noteWElem :: Node t => t -> Content
  noteWElem = wElem (noteType nt)
  notesElement = node notesWName ([xmlnsW], cs)
   where
    cs = separator nt : continuationSeparator nt : map note ns
  refAttrs ref = [ wIdAttr ref ]
  refR = r (noteRef nt)
  note (ref, []) =
    noteWElem ( refAttrs ref, p [ refR ] )
  note (ref, blocks@(Table' {}:_)) =
    noteWElem ( refAttrs ref, p $ refR : map fromBlock' blocks )
  note (ref, (Paragraph' mStyle props runs) : blocks) =
    noteWElem ( refAttrs ref, p' : map fromBlock' blocks )
   where
    p' = p $ fromParaProps mStyle props : refR : fromRuns' runs

separator :: Note -> Content
separator nt = note' nt 0 Separator
  [ p [ r [ wElem "separator" () ] ] ]

continuationSeparator :: Note -> Content
continuationSeparator nt = note' nt 1 ContinuationSeparator
  [ p [ r [ wElem "continuationSeparator" () ] ] ]

note' :: Note -> Int -> NoteType -> [Content] -> Content
note' nt ref fnType cs =
  wElem (noteType nt) ( wAttr "type" fnType' <> [wIdAttr ref], cs )
 where
  fnType' = case fnType of
    NormalNote        -> "normal"
    Separator             -> "separator"
    ContinuationSeparator -> "continuationSeparator"

noteRef :: Note -> [Content]
noteRef nt = case nt of
  FootNote -> footnoteRef
  EndNote  -> endnoteRef

footnoteRef :: [Content]
footnoteRef =
  [ rPr [rStyle "Footnote Reference"]
  , wElem "footnoteRef" ()
  ]

endnoteRef :: [Content]
endnoteRef =
  [ rPr [rStyle "Endnote Reference"]
  , wElem "endnoteRef" ()
  ]

-- |Create the header[1..n].xml or footer[1..n].xml entries.
marginalEntries :: Marginal -> [(a, [Block'])] -> [Integer -> Entry]
marginalEntries marginal marginals = snd $ mapAccumL marginalEntry 1 marginals
 where
  marginalEntry :: Int -> (a, [Block']) -> (Int, Integer -> Entry)
  marginalEntry n (_, blocks) = (n + 1, \st -> toEntry
    (runningMarginalXml marginal n)
    st
    (BS.pack $ ppTopElement marginalXml))
   where
    marginalXml =
      let cs = map fromBlock' blocks
      in  node (wName marginalName') ([xmlnsW], cs)
    marginalName' = case marginal of Header -> "hdr"; Footer -> "ftr"

marginalName :: Marginal -> String
marginalName marginal = case marginal of
                          Header -> "header"
                          Footer -> "footer"

runningMarginalXml :: Marginal -> Int -> String
runningMarginalXml marginal n = marginalName marginal <> show n <>".xml"

-- |The minimal [Content_Types].xml XML.
contentTypes :: Bool -> Int -> Int -> Bool -> Bool -> Element
contentTypes hasNumbering hRefCount fRefCount hasFootnotes hasEndnotes =
  node (unqual "Types") ([contentTypesAttr],
    [ default' "rels"
               "application/vnd.openxmlformats-package.relationships+xml"
    , default' "xml" "application/xml"
    , default' "png" "image/png"
    , override "document.xml" "document.main"
    , override "styles.xml" "styles"
    , override "settings.xml" "settings"
    ] <>
    [ override "numbering.xml" "numbering" | hasNumbering ] <>
    [ override "footnotes.xml" "footnotes" | hasFootnotes ] <>
    [ override "endnotes.xml" "endnotes" | hasEndnotes ] <>
    overrides Header hRefCount <>
    overrides Footer fRefCount)

contentTypesAttr :: Attr
contentTypesAttr = Attr (unqual "xmlns")
                        "http://schemas.openxmlformats.org/package/2006/content-types"

default' :: String -> String -> Content
default' extName contentType =
  uElem "Default" $ attr "Extension" extName <> attr "ContentType" contentType

override :: String -> String -> Content
override partName contentType =
  uElem "Override" $ partAttr <> contentTypeAttr
 where
  partAttr = attr "PartName" ("/" <> partName)
  contentTypeAttr = attr "ContentType"
    ("application/vnd.openxmlformats-officedocument.wordprocessingml." <>
       contentType <> "+xml")

overrides :: Marginal -> Int -> [Content]
overrides _ 0 = []
overrides marginal refCount = map fromRef [1 .. refCount]
 where
  fromRef :: Int -> Content
  fromRef n = override (runningMarginalXml marginal n) (marginalName marginal)

-- |The minimal .rels XML
rels :: Element
rels = fromMaybe undefined $ parseXMLDoc $
  "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" <>
  "<Relationship Id=\"rId1\" " <>
  "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" " <>
  "Target=\"document.xml\"/>" <>
  "</Relationships>"

-- |Create document.xml.rels XML
documentXmlRels :: Bool
                -> [(Int, [Block'])]
                -> [(Int, [Block'])]
                -> Bool
                -> Bool
                -> [(Int, DynamicImage)]
                -> Element
documentXmlRels
  hasNumbering headers footers hasFootnotes hasEndnotes drawings =
    node (unqual "Relationships") ([relsAttr],
      [ relationship' 1 "styles"
      , relationship' 2 "settings"
      ] <>
      [ relationship' 3 "footnotes" | hasFootnotes ] <>
      [ relationship' 4 "endnotes" | hasEndnotes ] <>
      [ relationship' 5 "numbering" | hasNumbering ] <>
      fromHeaders headers <>
      fromFooters footers <>
      fromDrawings drawings)
 where
  relationship' :: Int -> String -> Content
  relationship' n part = relationship
    ("rId" <> show n)
    ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/" <>
      part)
    (part <> ".xml")

fromHeaders :: [(Int, a)] -> [Content]
fromHeaders headers = snd $ mapAccumL fromHeader 0 headers

fromHeader:: Int -> (Int, a) -> (Int, Content)
fromHeader n (relId, _) = (n+1, relationship
  relId'
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
  (runningMarginalXml Header (n + 1)))
 where
  relId' = "rId" <> show relId

fromFooters :: [(Int, a)] -> [Content]
fromFooters footers = snd $ mapAccumL fromFooter 0 footers

fromFooter :: Int -> (Int, a) -> (Int, Content)
fromFooter n (relId, _) = (n + 1, relationship
  relId'
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
  (runningMarginalXml Footer (n + 1)))
 where
  relId' = "rId" <> show relId

fromDrawings :: [(Int, a)] -> [Content]
fromDrawings drawings = snd $ mapAccumL fromDrawing' 0 drawings

fromDrawing' :: Int -> (Int, a) -> (Int, Content)
fromDrawing' n (relId, _) = (n + 1, relationship
  relId'
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  ("media/image" <> show (n + 1) <> ".png"))
 where
  relId' = "picId" <> show relId

relsAttr :: Attr
relsAttr = Attr (unqual "xmlns")
                "http://schemas.openxmlformats.org/package/2006/relationships"

relationship :: String -> String -> String -> Content
relationship relId relType relTarget =
  uElem "Relationship" $
    attr "Id" relId <>
    attr "Type" relType <>
    attr "Target" relTarget

wName :: String -> QName
wName n = (unqual n) {qPrefix = Just "w"}

rName :: String -> QName
rName n = (unqual n) {qPrefix = Just "r"}

wpName :: String -> QName
wpName n = (unqual n) {qPrefix = Just "wp"}

picName :: String -> QName
picName n = (unqual n) {qPrefix = Just "pic"}

aName :: String -> QName
aName n = (unqual n) {qPrefix = Just "a"}

xmlns :: String -> String -> Attr
xmlns nsName = Attr (blank_name {qName = nsName, qPrefix = Just "xmlns"})

xmlnsW :: Attr
xmlnsW =
  xmlns "w"
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

xmlnsR :: Attr
xmlnsR =
  xmlns "r"
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

xmlnsWp :: Attr
xmlnsWp =
  xmlns "wp"
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

xmlnsPic :: Attr
xmlnsPic =
  xmlns "pic"
        "http://schemas.openxmlformats.org/drawingml/2006/picture"

xmlnsA :: Attr
xmlnsA =
  xmlns "a"
        "http://schemas.openxmlformats.org/drawingml/2006/main"

class Display a where

  display :: a -> String

instance Display [Char] where

  display = id

instance Display Int where

  display = show

instance Display Bool where

  display = show

instance Display Justification where

  display j = case j of
    Start      -> "start"
    End        -> "end"
    Center     -> "center"
    Both       -> "both"
    Distribute -> "distribute"

instance Display LevelSuffix where

  display ls = case ls of
    TabSuffix   -> "tab"
    SpaceSuffix -> "space"
    NoSuffix    -> "nothing"

attr :: Display a => String -> a -> [Attr]
attr k v = [Attr (unqual k) (display v)]

wAttr :: Display a => String -> a -> [Attr]
wAttr k v  = [Attr (wName k) (display v)]

wBoolAttr :: String -> Bool -> Attr
wBoolAttr n isSet = Attr (wName n) (if isSet then "true" else "false")

wValAttr :: Display a => a -> Attr
wValAttr v = Attr (wName "val") (display v)

wIdAttr :: Display a  => a -> Attr
wIdAttr ref = Attr (wName "id") (display ref)

uElem :: Node t => String -> t -> Content
uElem n x = Elem $ node (unqual n) x

wElem :: Node t => String -> t -> Content
wElem n x = Elem $ node (wName n) x

wpElem :: Node t => String -> t -> Content
wpElem n x = Elem $ node (wpName n) x

picElem :: Node t => String -> t -> Content
picElem n x = Elem $ node (picName n) x

aElem :: Node t => String -> t -> Content
aElem n x = Elem $ node (aName n) x

wValElem :: Display a => String -> a -> Content
wValElem n v = wElem n (wValAttr v)

wToggleProp :: String -> Toggle -> [Content]
wToggleProp n (First mP) = wBoolProp n mP

wBoolProp :: String -> Maybe Bool -> [Content]
wBoolProp n mP = case mP of
  Nothing    -> []
  Just True  -> [ wElem n () ]
  Just False -> [ wValElem n "false" ]

--------------------------------------------------------------------------------
-- The Document Settings part
--------------------------------------------------------------------------------

-- |Helper function to create settings.xml XML.
settingsXml :: DocProps -> Bool -> Bool -> Element
settingsXml props hasFootnotes hasEndnotes = settings $
  [ footnotePr fnCs | hasFootnotes ] <>
  [ endnotePr enCs | hasEndnotes ] <>
  maybe [] (pure . defaultTabStop) (dPrDefaultTabStop props)
 where
  fnCs = maybe [] (pure . fromNumberFormat) (dPrFnPrNumFmt props) <>
         specialFootnotes
  enCs = maybe [] (pure . fromNumberFormat) (dPrEnPrNumFmt props) <>
         specialEndnotes

-- |Create the top-level element
settings :: [Content] -> Element
settings cs = node (wName "settings") ([xmlnsW], cs)

footnotePr :: [Content] -> Content
footnotePr = wElem "footnotePr"

endnotePr :: [Content] -> Content
endnotePr = wElem "endnotePr"

defaultTabStop :: Twip -> Content
defaultTabStop = wValElem "defaultTabStop"

-- |Create the special footnotes list. This assumes that footnotes with id 0 and
-- 1 will be special footnotes and the only special footnotes.
specialFootnotes :: [Content]
specialFootnotes = specialNotes FootNote

-- |Create the special endnotes list. This assumes that endnotes with id 0 and
-- 1 will be special endnotes and the only special endnotes.
specialEndnotes :: [Content]
specialEndnotes = specialNotes EndNote

specialNotes :: Note -> [Content]
specialNotes nt =
  [ wElem (noteType nt) (wIdAttr "0")
  , wElem (noteType nt) (wIdAttr "1")
  ]

fromNumberFormat :: NumberFormat -> Content
fromNumberFormat nf = wValElem "numFmt" nf'
 where
  nf' = case nf of
    Aiueo                        -> "aiueo"
    AiueoFullWidth               -> "aiueoFullWidth"
    ArabicAbjad                  -> "arabicAbjad"
    ArabicAlpha                  -> "arabicAlpha"
    BahtText                     -> "bahtText"
    Bullet                       -> "bullet"
    CardinalText                 -> "cardinalText"
    Chicago                      -> "chicago"
    ChineseCounting              -> "chineseCounting"
    ChineseCountingThousand      -> "chineseCountingThousand"
    ChineseLegalSimplified       -> "chineseLegalSimplified"
    Chosung                      -> "chosung"
    Custom                       -> "custom"
    Decimal                      -> "decimal"
    DecimalEnclosedCircle        -> "decimalEnclosedCircle"
    DecimalEnclosedCircleChinese -> "decimalEnclosedCircleChinese"
    DecimalEnclosedFullstop      -> "decimalEnclosedFullstop"
    DecimalEnclosedParen         -> "decimalEnclosedParen"
    DecimalFullWidth             -> "decimalFullWidth"
    DecimalFullWidth2            -> "decimalFullWidth2"
    DecimalHalfWidth             -> "decimalHalfWidth"
    DecimalZero                  -> "decimalZero"
    DollarText                   -> "dollarText"
    Ganada                       -> "ganada"
    Hebrew1                      -> "hebrew1"
    Hebrew2                      -> "hebrew2"
    Hex                          -> "hex"
    HindiConsonants              -> "hindiConsonants"
    HindiCounting                -> "hindiCounting"
    HindiNumbers                 -> "hindiNumbers"
    HindiVowels                  -> "hindiVowels"
    IdeographDigital             -> "ideographDigital"
    IdeographEnclosedCircle      -> "ideographEnclosedCircle"
    IdeographLegalTraditional    -> "ideographLegalTraditional"
    IdeographTraditional         -> "ideographTraditional"
    IdeographZodiac              -> "ideographZodiac"
    IdeographZodiacTraditional   -> "ideographZodiacTraditional"
    Iroha                        -> "iroha"
    IrohaFullWidth               -> "irohaFullWidth"
    JapaneseCounting             -> "japaneseCounting"
    JapaneseDigitalTenThousand   -> "japaneseDigitalTenThousand"
    JapaneseLegal                -> "japaneseLegal"
    KoreanCounting               -> "koreanCounting"
    KoreanDigital                -> "koreanDigital"
    KoreanDigital2               -> "koreanDigital2"
    KoreanLegal                  -> "koreanLegal"
    LowerLetter                  -> "lowerLetter"
    LowerRoman                   -> "lowerRoman"
    NoNumberFormat               -> "none"
    NumberInDash                 -> "numberInDash"
    Ordinal                      -> "ordinal"
    OrdinalText                  -> "ordinalText"
    RussianLower                 -> "russianLower"
    RussianUpper                 -> "russianUpper"
    TaiwaneseCounting            -> "taiwaneseCounting"
    TaiwaneseCountingThousand    -> "taiwaneseCountingThousand"
    TaiwaneseDigital             -> "taiwaneseDigital"
    ThaiCounting                 -> "thaiCounting"
    ThaiLetters                  -> "thaiLetters"
    ThaiNumbers                  -> "thaiNumbers"
    UpperLetter                  -> "upperLetter"
    UpperRoman                   -> "upperRoman"
    VietnameseCounting           -> "vietnameseCounting"

--------------------------------------------------------------------------------
-- The numbering part
--------------------------------------------------------------------------------

numberingXml :: NumDefs -> Element
numberingXml (NumDefs aNums nums) = numbering $
  fromAbstractNums aNums <>
  fromNumberings nums

numbering :: [Content] -> Element
numbering cs = node (wName "numbering") ([xmlnsW], cs)

fromAbstractNums :: [AbstractNum] -> [Content]
fromAbstractNums = map fromAbstractNum

fromAbstractNum :: AbstractNum -> Content
fromAbstractNum (AbstractNum props nLvls) =
  abstractNum (anPrAbstractNumId props) $
    fromNumLevels nLvls

abstractNum :: Int -> [Content] -> Content
abstractNum n cs = wElem "abstractNum" (wAttr "abstractNumId" n, cs)

fromNumLevels :: [NumLevel] -> [Content]
fromNumLevels = map fromNumLevel

fromNumLevel :: NumLevel -> Content
fromNumLevel nLvl = lvl (nlIlvl nLvl) $
  maybe [] (pure . start) (nlStart nLvl) <>
  maybe [] (pure . fromNumberFormat) (nlNumFmt nLvl) <>
  maybe [] (pure . lvlRestart) (nlLvlRestart nLvl) <>
  maybe [] (pure . pStyle) (nlPStyle nLvl) <>
  isLgl (nlIsLgl nLvl) <>
  maybe [] (pure . suff) (nlSuff nLvl) <>
  maybe [] (pure . lvlText) (nlLvlText nLvl) <>
  maybe [] (pure . lvlJc) (nlLvlJc nLvl) <>
  maybe [] (pure . pPr . fromParaProps') (nlPPr nLvl) <>
  maybe [] (pure . fromRunProps Nothing) (nlRPr nLvl)

lvl :: Int -> [Content] -> Content
lvl n cs = wElem "lvl" (wAttr "ilvl" n, cs)

start :: Int -> Content
start = wValElem "start"

lvlRestart :: Int -> Content
lvlRestart = wValElem "lvlRestart"

isLgl :: Maybe Bool -> [Content]
isLgl = wBoolProp "isLgl"

suff :: LevelSuffix -> Content
suff = wValElem "suff"

lvlText :: String -> Content
lvlText = wValElem "lvlText"

lvlJc :: Justification -> Content
lvlJc = wValElem "lvlJc"

fromNumberings :: [Numbering] -> [Content]
fromNumberings = map fromNumbering

fromNumbering :: Numbering -> Content
fromNumbering (Numbering props nLvls) = num (nPrNumId props) $
  wValElem "abstractNumId" (nPrAbstractNumId props) :
  fromLvlOverrides nLvls

num :: Int -> [Content] -> Content
num n cs = wElem "num" (wAttr "numId" n, cs)

fromLvlOverrides :: [NumLevel] -> [Content]
fromLvlOverrides = map fromLvlOverride

fromLvlOverride :: NumLevel -> Content
fromLvlOverride nLvl = lvlOverride (nlIlvl nLvl) [fromNumLevel nLvl]

lvlOverride :: Int -> [Content] -> Content
lvlOverride n cs = wElem "lvlOverride" (wAttr "ilvl" n, cs)

--------------------------------------------------------------------------------
-- The styles part
--------------------------------------------------------------------------------

-- |Helper function to create styles.xml XML.
stylesXml :: Styles -> Element
stylesXml (Styles mPprops mRprops ss) = styles $
  docDefaults mPprops mRprops : map fromStyle (toList ss)

-- |Create the top-level element
styles :: [Content] -> Element
styles cs = node (wName "styles") ([xmlnsW], cs)

docDefaults :: Maybe ParaProps -> Maybe RunProps -> Content
docDefaults mPprops mRprops = wElem "docDefaults" $
  maybe [] (\pProps -> [pPrDefault $ fromParaProps' pProps]) mPprops <>
  maybe [] (\rProps -> [rPrDefault [fromRunProps Nothing rProps]]) mRprops

pPrDefault :: [Content] -> Content
pPrDefault = wElem "pPrDefault"

rPrDefault :: [Content] -> Content
rPrDefault = wElem "rPrDefault"

fromStyle :: (String, Style) -> Content
fromStyle (n, ParaStyle mStyle mPprops mRprops isD isQf) =
  paragraphStyle n mStyle mPprops mRprops isD isQf
fromStyle (n, RunStyle mStyle mRprops isD isQf) =
  runStyle n mStyle mRprops isD isQf

paragraphStyle :: String
               -> Maybe String
               -> Maybe ParaProps
               -> Maybe RunProps
               -> Bool
               -> Bool
               -> Content
paragraphStyle styleName mStyleName mPprops mRprops isD isQf =
  wElem "style" ( wAttr "type" "paragraph" <>
                  wAttr "styleId" styleName <>
                  [ wBoolAttr "default" isD ]
                , cs )
 where
  cs = name styleName :
       maybe [] (pure . basedOn) mStyleName <>
       maybe [] (pure . pPr . fromParaProps') mPprops <>
       maybe [] (pure . fromRunProps Nothing) mRprops <>
       [ qFormat | isQf ]

name :: String -> Content
name = wValElem "name"

basedOn :: String -> Content
basedOn = wValElem "basedOn"

qFormat :: Content
qFormat = Elem $ node (wName "qFormat") ()

runStyle :: String
         -> Maybe String
         -> Maybe RunProps
         -> Bool
         -> Bool
         -> Content
runStyle styleName mStyleName mRprops isD isQf =
  wElem "style" ( wAttr "type" "character" <>
                  wAttr "styleId" styleName <>
                  [ wBoolAttr "default" isD ]
                , cs )
 where
  cs = name styleName :
       maybe [] (pure . basedOn) mStyleName <>
       maybe [] (pure . fromRunProps Nothing) mRprops <>
       [ qFormat | isQf ]

--------------------------------------------------------------------------------
-- The document part
--------------------------------------------------------------------------------

-- |Helper function to create minimal document.xml XML.
documentXml :: Docx
            -> (Element, OtherParts)
documentXml d = runState (do
  c <- section $ docxSections d
  document [body c]) emptyOtherParts

document :: [Content] -> State OtherParts Element
document cs = pure $ node (wName "document") (xmlnss, cs)
 where
  xmlnss = [xmlnsW, xmlnsR, xmlnsPic, xmlnsA, xmlnsWp]

body :: [Content] -> Content
body = wElem "body"

section :: [Section] -> State OtherParts [Content]
section [] = pure []
section [Section sp bs] = contents True sp bs
section ((Section sp bs):ss) = do
  cs <- contents False sp bs
  cs' <- section ss
  pure $ cs <> cs'

contents :: Bool
         -> SectionProps
         -> [Block]
         -> State OtherParts [Content]
contents True sp [] = fromSectionProps sp
contents False _ [] = undefined
contents False sp [block] = fromBlockWSp sp block
contents isLastSection sp (block:bs) = do
  cs <- fromBlock block
  cs' <- contents isLastSection sp bs
  pure $ cs <> cs'

fromSectionProps :: SectionProps
                 -> State OtherParts [Content]
fromSectionProps sp = do
  cs <- concatMapM marginalRef (maybeAssocs $ sPrMarginals sp)
  pure $ pure $ sectPr $
    [ pgSz $ sPrPgSz sp
    , pgMar $ sPrPgMar sp ] <>
    cs <>
    maybe [] (pure . footnotePr . pure .fromNumberFormat) (sPrFnPrNumFmt sp) <>
    maybe [] (pure . endnotePr . pure . fromNumberFormat) (sPrEnPrNumFmt sp)

marginalRef :: (Marginal, MarginalType, [Block'])
            -> State OtherParts [Content]
marginalRef (_, _, []) = pure []
marginalRef (m, mt, blocks) = do
  headerFooterNo <- gets opHeaderFooterNo
  marginals <- gets (case m of Header -> opHeaders; Footer -> opFooters)
  let headerFooterNo' = headerFooterNo + 1
      marginals' = (headerFooterNo', blocks) : marginals
      ref' = "rId" <> show headerFooterNo'
  modify (case m of
        Header -> \s -> s { opHeaderFooterNo = headerFooterNo'
                     , opHeaders = marginals' }
        Footer -> \s -> s { opHeaderFooterNo = headerFooterNo'
                     , opFooters = marginals' })
  pure $ pure $ wElem refName $
    [ Attr (QName "id" Nothing (Just "r")) ref' ] <>
    wAttr "type" refType
 where
  refName = case m of Header -> "headerReference"; Footer -> "footerReference"
  refType = case mt of
              DefaultMarginal -> "default"
              FirstMarginal   -> "first"
              EvenMarginal    -> "even"

maybeAssocs :: Array (Marginal, MarginalType) [Block']
            -> [(Marginal, MarginalType, [Block'])]
maybeAssocs a = concatMap maybeAssocs' (assocs a)

maybeAssocs' :: ((Marginal, MarginalType), [Block'])
             -> [(Marginal, MarginalType, [Block'])]
maybeAssocs' (_, []) = []
maybeAssocs' ((m, mt), blocks) = [(m, mt, blocks)]

sectPr :: [Content] -> Content
sectPr = wElem "sectPr"

pgSz :: PageSize -> Content
pgSz (w, h) = wElem "pgSz" $ wAttr "w" w <> wAttr "h" h

pgMar :: PageMargins -> Content
pgMar (PageMargins to bo tm bm lm rm h f g) = wElem "pgMar" $
  wAttr "top" (if to then -tm else tm) <>
  wAttr "bottom" (if bo then -bm else bm) <>
  wAttr "left" lm <>
  wAttr "right" rm <>
  wAttr "header" h <>
  wAttr "footer" f <>
  wAttr "gutter" g

fromBlockWSp :: SectionProps
             -> Block
             -> State OtherParts [Content]
fromBlockWSp sp block = case block of
  Paragraph mStyle props runs -> do
    c <- fromParagraphWSp sp mStyle props runs
    pure [c]
  Table props tGrid rows -> do
    c <- fromTable props tGrid rows
    c' <- emptyPara
    pure $ c : [c']
 where
  emptyPara = fromParagraphWSp sp Nothing defaultParaProps []

fromBlock :: Block -> State OtherParts [Content]
fromBlock block = case block of
  Paragraph mStyle props runs -> do
    c <- fromParagraph mStyle props runs
    pure [c]
  Table props tGrid rows      -> do
    c <- fromTable props tGrid rows
    pure [c]

fromBlock' :: Block' -> Content
fromBlock' block = case block of
  Paragraph' mStyle props runs -> fromParagraph' mStyle props runs
  Table' props tGrid rows      -> fromTable' props tGrid rows

--------------------------------------------------------------------------------
-- Paragraph blocks
--------------------------------------------------------------------------------

fromParagraphWSp :: SectionProps
                 -> Maybe String
                 -> ParaProps
                 -> [Run]
                 -> State OtherParts Content
fromParagraphWSp sp mStyle props runs = do
  cs <- fromRuns runs
  cs' <- fromParaPropsWSp sp mStyle props
  pure $ p $ cs' : cs

fromParagraph :: Maybe String -> ParaProps -> [Run] -> State OtherParts Content
fromParagraph mStyle props runs = do
  cs <- fromRuns runs
  pure $ p $ fromParaProps mStyle props : cs

fromParagraph' :: Maybe String -> ParaProps -> [Run'] -> Content
fromParagraph' mStyle props runs =
  p $ fromParaProps mStyle props : fromRuns' runs

p :: [Content] -> Content
p = wElem "p"

fromParaPropsWSp :: SectionProps
                 -> Maybe String
                 -> ParaProps
                 -> State OtherParts Content
fromParaPropsWSp sp mStyle props = do
  cs <- fromSectionProps sp
  pure $ pPr $
    cs <>
    maybe [] (pure . pStyle) mStyle <>
    fromParaProps' props

fromParaProps :: Maybe String
              -> ParaProps
              -> Content
fromParaProps mStyle props = pPr $
  maybe [] (pure . pStyle) mStyle <>
  fromParaProps' props

pPr :: [Content] -> Content
pPr = wElem "pPr"

pStyle :: String -> Content
pStyle = wValElem "pStyle"

fromParaProps' :: ParaProps -> [Content]
fromParaProps' pProps =
  maybe [] fromNumPr (pPrNumPr pProps) <>
  fromTabs (pPrTabs pProps) <>
  maybe [] fromSpacing (pPrSpacing pProps) <>
  maybe [] (pure . ind) (pPrInd pProps) <>
  maybe [] (pure . jc) (pPrJc pProps)

fromNumPr :: NumProps -> [Content]
fromNumPr np = pure . numPr $
  maybe [] (pure . ilvl) (numPrIlvl np) <>
  maybe [] (pure . numId) (numPrNumId np)

numPr :: [Content] -> Content
numPr = wElem "numPr"

ilvl :: Int -> Content
ilvl = wValElem "ilvl"

numId :: Int -> Content
numId = wValElem "numId"

fromSpacing :: Spacing -> [Content]
fromSpacing spacing = [ wElem "spacing" sp ]
 where
  sp = maybe [] (wAttr "before") (spacingBefore spacing) <>
       maybe [] (wAttr "after") (spacingAfter spacing) <>
       maybe [] (\(lr, x) -> fromLineRule lr <> wAttr "line" x)
         (spacingLine spacing)

fromLineRule :: LineRule -> [Attr]
fromLineRule lr = wAttr "lineRule" $ case lr of
                                       Auto    -> "auto"
                                       AtLeast -> "atLeast"
                                       Exactly -> "exactly"

fromTabs :: [Tab] -> [Content]
fromTabs [] = []
fromTabs ts = [tabs $ map fromTab ts]

tabs :: [Content] -> Content
tabs = wElem "tabs"

fromTab :: Tab -> Content
fromTab (ts, x) = tab ts x

tab :: TabStyle -> Int -> Content
tab ts x = wElem "tab" $ [ fromTabStyle ts ] <> wAttr "pos" x

fromTabStyle :: TabStyle -> Attr
fromTabStyle ts = wValAttr ts'
 where
  ts' = case ts of
          BarTab     -> "bar"
          ClearTab   -> "clear"
          EndTab     -> "end"
          StartTab   -> "start"
          CenterTab  -> "center"
          DecimalTab -> "decimal"

jc :: Justification -> Content
jc j = wElem "jc" (fromJustification j)

fromJustification :: Justification -> Attr
fromJustification = wValAttr

-- |Reference: /ECMA-376-1:2016, Fundamentals and Markup Language Reference
-- ?? 17.3.1.12/.
ind :: Indentation -> Content
ind i = wElem "ind" (fromIndentation i)

fromIndentation :: Indentation -> [Attr]
fromIndentation i =
  maybe [] (wAttr "start") (indStart i) <>
  maybe [] (wAttr "end") (indEnd i) <>
  maybe [] (wAttr "hanging") (indHanging i) <>
  maybe [] (wAttr "firstLine") (indFirstLine i)

--------------------------------------------------------------------------------
-- Runs
--------------------------------------------------------------------------------

fromRuns :: [Run] -> State OtherParts [Content]
fromRuns = mapM fromRun

fromRun :: Run -> State OtherParts Content
fromRun (Run mStyle props rcs) = do
  cs <- mapM fromRunContent rcs
  pure $ r $ fromRunProps mStyle props : cs
fromRun (Footnote mStyle props blocks) = do
  footnoteNo <- gets opFootnoteNo
  footnotes  <- gets opFootnotes
  let footnoteNo' = footnoteNo + 1
      footnotes' = (footnoteNo', blocks) : footnotes
  modify (\s -> s { opFootnoteNo = footnoteNo', opFootnotes = footnotes' })
  pure $ footnoteReference mStyle props footnoteNo'
fromRun (Endnote mStyle props blocks) = do
  endnoteNo <- gets opEndnoteNo
  endnotes  <- gets opEndnotes
  let endnoteNo' = endnoteNo + 1
      endnotes' = (endnoteNo', blocks) : endnotes
  modify (\s -> s { opEndnoteNo = endnoteNo', opEndnotes = endnotes' })
  pure $ endnoteReference mStyle props endnoteNo'

fromRuns' :: [Run'] -> [Content]
fromRuns' = map fromRun'

fromRun' :: Run' -> Content
fromRun' (Run' mStyle props rcs) =
  let cs = evalState (mapM fromRunContent rcs) emptyOtherParts
  in  r $ fromRunProps mStyle props : cs

r :: [Content] -> Content
r = wElem "r"

fromRunProps :: Maybe String -> RunProps -> Content
fromRunProps mStyle props = rPr $
  maybe [] (pure . rStyle) mStyle <>
  concatMap (\f -> f props)
    [ toWToggleProp "b" rPrIsBold
    , toWToggleProp "i" rPrIsItalic
    , toProp u rPrU
    , toWToggleProp "strike" rPrIsStrike
    , toWToggleProp "dStrike" rPrIsDStrike
    , toWToggleProp "caps" rPrIsCaps
    , toWToggleProp "smallCaps" rPrIsSmallCaps
    , toWToggleProp "emboss" rPrIsEmboss
    , toWToggleProp "imprint" rPrIsImprint
    , toWToggleProp "outline" rPrIsOutline
    , toWToggleProp "shadow" rPrIsShadow
    , toProp color rPrColor
    , toProp rFonts rPrRFonts
    , toProp vertAlign rPrVertAlign
    , toProp sz rPrSz
    ]

toProp :: (a -> Content) -> (b -> Last a) -> b -> [Content]
toProp tag getter props = maybe [] (pure . tag) (getLast $ getter props)

toWToggleProp :: String -> (a -> Toggle) -> a -> [Content]
toWToggleProp n getter props = wToggleProp n (getter props)

rPr :: [Content] -> Content
rPr = wElem "rPr"

rStyle :: String -> Content
rStyle = wValElem "rStyle"

u :: Underline -> Content
u ul = wElem "u" (fromUnderline ul)

fromUnderline :: Underline -> [Attr]
fromUnderline (Underline mColor up) =
  maybe [] (fromColor "color") mColor <>
  [fromUnderlinePattern up]

fromColor :: String -> RGB Word8 -> [Attr]
fromColor n (RGB red green blue) = wAttr n rgb'
 where
  rgb' = showHex' red $ showHex' green $ showHex' blue ""
  showHex' x = if x < 16
    then ("0" <>) . showHex x
    else showHex x

fromUnderlinePattern :: UnderlinePattern -> Attr
fromUnderlinePattern up = wValAttr up'
 where
  up' = case up of
          Dash            -> "dash"
          DashDotDotHeavy -> "DashDotDotHeavy"
          DashDotHeavy    -> "dashDotHeavy"
          DashedHeavy     -> "dashedHeavy"
          DashLong        -> "dashLong"
          DashLongHeavy   -> "dashLongHeavy"
          DotDash         -> "dotDash"
          DotDotDash      -> "dotDotDash"
          Dotted          -> "dotted"
          DottedHeavy     -> "dottedHeavy"
          Double          -> "double"
          NoUnderline     -> "none"
          Single          -> "single"
          Thick           -> "thick"
          Wave            -> "wave"
          WavyDouble      -> "wavyDouble"
          WavyHeavy       -> "wavyHeavy"
          Words           -> "words"

color :: RGB Word8 -> Content
color rgb = wElem "color" (fromColor "val" rgb)

rFonts :: Fonts -> Content
rFonts fs = wElem "rFonts" (fromFonts fs)

fromFonts :: Fonts -> [Attr]
fromFonts fs =
  maybe [] (wAttr "ascii") (rFontsAscii fs) <>
  maybe [] (wAttr "cs") (rFontsCs fs) <>
  maybe [] (wAttr "eastAsia") (rFontsEastAsia fs) <>
  maybe [] (wAttr "hAnsi") (rFontsHAnsi fs)

sz :: Int -> Content
sz = wValElem "sz"

vertAlign :: VertAlign -> Content
vertAlign va = wElem "vertAlign" (fromVertAlign va)

fromVertAlign :: VertAlign -> Attr
fromVertAlign va = wValAttr va'
 where
  va' = case va of
          Baseline    -> "baseline"
          Subscript   -> "subscript"
          Superscript -> "superscript"

fromRunContent :: RunContent -> State OtherParts Content
fromRunContent (RunText text) = pure $ t text
fromRunContent (Break bt) = pure $ br bt
fromRunContent NoBreakHyphen = pure noBreakHyphen
fromRunContent SoftHyphen = pure softHyphen
fromRunContent (Symbol fontName code) = pure $ sym fontName code
fromRunContent Tab = pure tab'
fromRunContent (Drawing d) = fromDrawing d

footnoteReference :: Maybe StyleName -> RunProps -> Int -> Content
footnoteReference mStyle props ref =
  r
    [ fromRunProps mStyle props
    , wElem "footnoteReference" (wIdAttr ref)
    ]

endnoteReference :: Maybe StyleName -> RunProps -> Int -> Content
endnoteReference mStyle props ref =
  r
    [ fromRunProps mStyle props
    , wElem "endnoteReference" (wIdAttr ref)
    ]

t :: String -> Content
t text = wElem "t" (Attr xmlSpace "preserve", text)
 where
  xmlSpace = QName "space" Nothing (Just "xml")

br :: BreakType -> Content
br (TextWrapping ct) = wElem "br" $ wAttr "type" "column" <>
                                    fromClearType ct
br ColumnBreak = wElem "br" (wAttr "type" "column")
br PageBreak = wElem "br" (wAttr "type" "page")

fromClearType :: ClearType -> [Attr]
fromClearType ct = wAttr "clear" ct'
 where
  ct' = case ct of
          NoneClear  -> "none"
          LeftClear  -> "left"
          RightClear -> "right"
          AllClear   -> "all"

noBreakHyphen :: Content
noBreakHyphen = wElem "noBreakHyphen" ()

softHyphen :: Content
softHyphen = wElem "softHyphen" ()

sym :: String -> Int -> Content
sym fontName code = wElem "sym" $ wAttr "font" fontName <>
                                  wAttr "char" (showHex code [])

tab' :: Content
tab' = wElem "tab" ()

fromDrawing :: DynamicImage -> State OtherParts Content
fromDrawing d = do
  drawingNo <- gets opDrawingNo
  drawings  <- gets opDrawings
  let drawingNo' = drawingNo + 1
      drawings' = (drawingNo', d) : drawings
      w = pxToEmu $ dynamicMap imageWidth d
      h = pxToEmu $ dynamicMap imageHeight d
  modify (\s -> s {opDrawingNo = drawingNo', opDrawings = drawings'})
  pure $ drawing
    [ inline
        [ extent w h
        , docPr drawingNo'
        , graphic
            [ graphicData
                [ pic
                    [ nvPicPr [ cNvPr drawingNo', cNvPicPr, nvPr ]
                    , blipFill [blip drawingNo']
                    , spPr
                        [ xfrm [ off 0 0, ext w h ]
                        , prstGeom
                        ]
                    ]
                ]
            ]
        ]
    ]

drawing :: [Content] -> Content
drawing = wElem "drawing"

inline :: [Content] -> Content
inline = wpElem "inline"

extent :: Int -> Int -> Content
extent w h = wpElem "extent" [ Attr (unqual "cx") (show w)
                             , Attr (unqual "cy") (show h)
                             ]

docPr :: Int -> Content
docPr n = wpElem "docPr" [ Attr (unqual "id") n'
                         , Attr (unqual "name") ("image" <> n' <> ".png")
                         ]
 where
  n' = show n

graphic :: [Content] -> Content
graphic = aElem "graphic"

graphicData :: [Content] -> Content
graphicData cs = aElem "graphicData" ([Attr (unqual "uri") "http://schemas.openxmlformats.org/drawingml/2006/picture"], cs)

pic :: [Content] -> Content
pic = picElem "pic"

nvPicPr :: [Content] -> Content
nvPicPr = picElem "nvPicPr"

cNvPr :: Int -> Content
cNvPr n = picElem "cNvPr" [ Attr (unqual "id") n'
                             , Attr (unqual "name") ("image" <> n' <> ".png")
                             ]
 where
  n' = show n

cNvPicPr :: Content
cNvPicPr = picElem "cNvPicPr" ()

nvPr :: Content
nvPr = picElem "nvPr" ()

blipFill :: [Content] -> Content
blipFill = picElem "blipFill"

blip :: Int -> Content
blip ref = aElem "blip" (Attr (rName "embed") ("picId" <> show ref))

spPr :: [Content] -> Content
spPr = picElem "spPr"

xfrm :: [Content] -> Content
xfrm = aElem "xfrm"

off :: Int -> Int -> Content
off x y = aElem "off" [ Attr (unqual "x") (show x)
                      , Attr (unqual "y") (show y)
                      ]

ext :: Int -> Int -> Content
ext cx cy = aElem "ext" [ Attr (unqual "cx") (show cx)
                        , Attr (unqual "cy") (show cy)
                        ]

prstGeom :: Content
prstGeom = aElem "prstGeom" (Attr (unqual "prst") "rect")

--------------------------------------------------------------------------------
-- Table blocks
--------------------------------------------------------------------------------

fromTable :: TableProps -> TableGrid -> [Row] -> State OtherParts Content
fromTable tProps tGrid tRows = do
  cs <- fromRows tRows
  pure $ tbl $ fromTableProps tProps : fromTableGrid tGrid : cs

fromRows :: [Row] -> State OtherParts [Content]
fromRows = mapM fromRow

fromRow :: Row -> State OtherParts Content
fromRow (Row rProps cells) = do
  cs <- fromCells cells
  pure $ tr $ fromRowProps rProps : cs

fromTable' :: TableProps -> TableGrid -> [Row'] -> Content
fromTable' tProps tGrid tRows =
  tbl $ fromTableProps tProps : fromTableGrid tGrid : fromRows' tRows

fromRows' :: [Row'] -> [Content]
fromRows' = map fromRow'

fromRow' :: Row' -> Content
fromRow' (Row' rProps cells) =
  tr $ fromRowProps rProps : fromCells' cells

tbl :: [Content] -> Content
tbl = wElem "tbl"

fromTableProps :: TableProps -> Content
fromTableProps tProps = tblPr $
  maybe [] tblW (tblPrTblW tProps) <>
  maybe [] tblLayout (tblPrTblLayout tProps) <>
  maybe [] fromTblCellMar (tblPrCellMar tProps)

tblPr :: [Content] -> Content
tblPr = wElem "tblPr"

tblW :: Twip -> [Content]
tblW w = pure $ wElem "tblW" $ wAttr "w" w <> wAttr "type" "dxa"

tblLayout :: TableLayout -> [Content]
tblLayout tl = pure $ wElem "tblLayout" (fromTableLayout tl)

fromTableLayout :: TableLayout -> [Attr]
fromTableLayout tl = wAttr "type" $ case tl of
                                      AutoLayout  -> "auto"
                                      FixedLayout -> "fixed"

fromTblCellMar :: CellMargins -> [Content]
fromTblCellMar cms = pure $ tblCellMar $
  maybe [] topMargin (tcMarginW <$> cmTop cms) <>
  maybe [] bottomMargin (tcMarginW <$> cmBottom cms) <>
  maybe [] startMargin (tcMarginW <$> cmStart cms) <>
  maybe [] endMargin (tcMarginW <$> cmEnd cms)

tblCellMar :: [Content] -> Content
tblCellMar = wElem "tblCellMar"

topMargin :: Twip -> [Content]
topMargin = cellMargin "top"

bottomMargin :: Twip -> [Content]
bottomMargin = cellMargin "bottom"

startMargin :: Twip -> [Content]
startMargin = cellMargin "start"

endMargin :: Twip -> [Content]
endMargin = cellMargin "end"

cellMargin :: String -> Twip -> [Content]
cellMargin n w = pure $ wElem n $ wAttr "w" w <> wAttr "type" "dxa"

fromTableGrid :: TableGrid -> Content
fromTableGrid tGrid = tblGrid $ map fromGridCol tGrid

tblGrid :: [Content] -> Content
tblGrid = wElem "tblGrid"

fromGridCol :: GridCol -> Content
fromGridCol = gridCol

gridCol :: Twip -> Content
gridCol w = wElem "gridCol" $ wAttr "w" w

tr :: [Content] -> Content
tr = wElem "tr"

fromRowProps :: RowProps -> Content
fromRowProps rProps = trPr $
  wBoolProp "cantSplit" (trPrCantSplit rProps) <>
  wBoolProp "tblHeader" (trPrTblHeader rProps)

trPr :: [Content] -> Content
trPr = wElem "trPr"

fromCells :: [Cell] -> State OtherParts [Content]
fromCells = mapM fromCell

fromCell :: Cell -> State OtherParts Content
fromCell (Cell cProps content) = do
  cs <- concatMapM fromBlock content
  pure $ tc $ [ tcPr $ maybe [] fromCellBorders mCbs <> [tcW w] ] <>
    cs
 where
  w = tcPrTcW cProps
  mCbs = tcPrTcBorders cProps

fromCells' :: [Cell'] -> [Content]
fromCells' = map fromCell'

fromCell' :: Cell' -> Content
fromCell' (Cell' cProps content) =
  tc $ [ tcPr $ maybe [] fromCellBorders mCbs <> [tcW w] ] <>
    map fromBlock' content
 where
  w = tcPrTcW cProps
  mCbs = tcPrTcBorders cProps

tc :: [Content] -> Content
tc = wElem "tc"

tcPr :: [Content] -> Content
tcPr = wElem "tcPr"

tcW :: Twip -> Content
tcW w = wElem "tcW" $ wAttr "w" w <> wAttr "type" "dxa"

fromCellBorders :: CellBorders -> [Content]
fromCellBorders cbs = pure $ tcBorders $
  maybe [] top (tcBorderSz <$> cbTop cbs) <>
  maybe [] bottom (tcBorderSz <$> cbBottom cbs) <>
  maybe [] left (tcBorderSz <$> cbLeft cbs) <>
  maybe [] right (tcBorderSz <$> cbRight cbs)

tcBorders :: [Content] -> Content
tcBorders = wElem "tcBorders"

top :: EighthPt -> [Content]
top = cellBorder "top"

bottom :: EighthPt -> [Content]
bottom = cellBorder "bottom"

left :: EighthPt -> [Content]
left = cellBorder "left"

right :: EighthPt -> [Content]
right = cellBorder "right"

cellBorder :: String -> EighthPt -> [Content]
cellBorder n s = pure $ wElem n $ wAttr "sz" s <> [ wValAttr "single" ]
