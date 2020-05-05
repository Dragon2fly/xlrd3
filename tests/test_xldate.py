#!/usr/bin/env python
# Author:  mozman <mozman@gmx.at>
# Purpose: test xldate.py
# Created: 04.12.2010
# Copyright (C) 2010, Manfred Moitzi
# License: BSD licence
# Updated: 2020/05/06, Duc Tin
import pytest

from xlrd import xldate

DATEMODE = 0  # 1900-based


def test_date_as_tuple():
    date = xldate.xldate_as_tuple(2741., DATEMODE)
    assert date == (1907, 7, 3, 0, 0, 0)
    date = xldate.xldate_as_tuple(38406., DATEMODE)
    assert date == (2005, 2, 23, 0, 0, 0)
    date = xldate.xldate_as_tuple(32266., DATEMODE)
    assert date == (1988, 5, 3, 0, 0, 0)


def test_time_as_tuple():
    time = xldate.xldate_as_tuple(.273611, DATEMODE)
    assert time == (0, 0, 0, 6, 34, 0)
    time = xldate.xldate_as_tuple(.538889, DATEMODE)
    assert time == (0, 0, 0, 12, 56, 0)
    time = xldate.xldate_as_tuple(.741123, DATEMODE)
    assert time == (0, 0, 0, 17, 47, 13)


def test_xldate_from_date_tuple():
    date = xldate.xldate_from_date_tuple((1907, 7, 3), DATEMODE)
    assert date == pytest.approx(2741.)
    date = xldate.xldate_from_date_tuple((2005, 2, 23), DATEMODE)
    assert date == pytest.approx(38406.)
    date = xldate.xldate_from_date_tuple((1988, 5, 3), DATEMODE)
    assert date == pytest.approx(32266.)


def test_xldate_from_time_tuple():
    time = xldate.xldate_from_time_tuple((6, 34, 0))
    assert time == pytest.approx(.273611)
    time = xldate.xldate_from_time_tuple((12, 56, 0))
    assert time == pytest.approx(.538889)
    time = xldate.xldate_from_time_tuple((17, 47, 13))
    assert time == pytest.approx(.741123)


def test_xldate_from_datetime_tuple():
    date = xldate.xldate_from_datetime_tuple((1907, 7, 3, 6, 34, 0), DATEMODE)
    assert date == pytest.approx(2741.273611)
    date = xldate.xldate_from_datetime_tuple((2005, 2, 23, 12, 56, 0), DATEMODE)
    assert date == pytest.approx(38406.538889)
    date = xldate.xldate_from_datetime_tuple((1988, 5, 3, 17, 47, 13), DATEMODE)
    assert date == pytest.approx(32266.741123)
