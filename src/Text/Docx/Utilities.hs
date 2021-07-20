{-|
Module      : Text.Docx.Utilities
Description : Helper functions
Copyright   : (c) Mike Pilgrem, 2021
License     : BSD-3
Maintainer  : public@pilgrem.com
Stability   : experimental
Portability : POSIX, Windows

-}
module Text.Docx.Utilities
  ( cmToTwip
  , ptToTwip
  , ptToHalfPt
  , toggle
  ) where

import Text.Docx.Types

cmToTwip :: Double -> Twip
cmToTwip cm = round (cm * 72.0 * 20.0 / 2.54)

ptToTwip :: Double -> Twip
ptToTwip pt = round (pt * 20.0)

ptToHalfPt :: Double -> Int
ptToHalfPt pt = round (pt * 2.0)

toggle :: Bool -> Toggle
toggle = pure
