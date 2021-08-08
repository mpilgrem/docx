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
  , cmToEmu
  , pxToEmu
  , ptToHalfPt
  , toggle
  ) where

import Text.Docx.Types

-- |Yields the twip (twentieth of a point) equivalent (rounded) of centimetres.
cmToTwip :: Double -> Twip
cmToTwip cm = round (cm * 72.0 * 20.0 / 2.54)

-- |Yields the twip (twentieth of a point) equivalent (rounded) of points.
ptToTwip :: Double -> Twip
ptToTwip pt = round (pt * 20.0)

-- |Yields the EMU (360,000th of a centimetre) equivalent (rounded) of
-- centimetres.
cmToEmu :: Double -> Emu
cmToEmu cm = round (cm * 360000.0)

pxToEmu :: Int -> Int
pxToEmu px = px * 9525

-- |Yields the half-point equivalent (rounded) of points.
ptToHalfPt :: Double -> HalfPt
ptToHalfPt pt = round (pt * 2.0)

-- |Yields the toggle setting equivalent of a boolean value.
toggle :: Bool -> Toggle
toggle = pure
