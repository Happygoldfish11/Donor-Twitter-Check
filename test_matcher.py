"""
Test suite for the Twitter Finder application.

Run with:  pytest tests/ -v
"""

import pytest
import io
import pandas as pd
from unittest.mock import patch, MagicMock

import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from src.matcher import (
    Person, TwitterProfile, MatchResult,
    score_match, _normalise, _build_queries,
    _parse_twitter_url, build_profile_from_candidate,
)
from src.spreadsheet import load_persons, write_results


# ---------------------------------------------------------------------------
# Person model
# ---------------------------------------------------------------------------

class TestPerson:
    def test_full_name(self):
        p = Person("John", "Smith")
        assert p.full_name == "John Smith"

    def test_full_name_strips_whitespace(self):
        p = Person("  Jane  ", "  Doe  ")
        assert p.full_name == "Jane Doe"

    def test_city_extraction_comma_format(self):
        p = Person("A", "B", address="123 Main St, Brooklyn, NY 10001")
        assert p.city == "Brooklyn"

    def test_city_extraction_single_part(self):
        p = Person("A", "B", address="Chicago")
        assert p.city == "Chicago"

    def test_state_extraction(self):
        p = Person("A", "B", address="123 Main St, Austin, TX 78701")
        assert p.state == "TX"

    def test_state_not_found(self):
        p = Person("A", "B", address="London, UK")
        assert p.state == ""

    def test_empty_address(self):
        p = Person("A", "B")
        assert p.city == ""
        assert p.state == ""


# ---------------------------------------------------------------------------
# Score matching
# ---------------------------------------------------------------------------

class TestScoreMatch:
    def _profile(self, **kwargs):
        defaults = dict(
            handle="@johndoe",
            display_name="John Doe",
            bio="Software engineer at Acme Corp in New York",
            location="New York, NY",
            followers=500,
            verified=False,
            created_year=2018,
            profile_url="https://twitter.com/johndoe",
        )
        defaults.update(kwargs)
        return TwitterProfile(**defaults)

    def test_exact_name_match_scores_high(self):
        person = Person("John", "Doe", address="New York, NY", employer="Acme Corp")
        profile = self._profile()
        score, breakdown = score_match(person, profile)
        assert score >= 75
        assert "name_exact" in breakdown or "name_both_parts" in breakdown

    def test_employer_match_adds_points(self):
        person = Person("John", "Doe", employer="Acme Corp")
        profile = self._profile()
        score, breakdown = score_match(person, profile)
        assert "employer_match" in breakdown or "employer_partial" in breakdown

    def test_city_match_adds_points(self):
        person = Person("John", "Doe", address="New York, NY")
        profile = self._profile()
        score, breakdown = score_match(person, profile)
        assert "city_match" in breakdown

    def test_state_match_adds_points(self):
        person = Person("John", "Doe", address="Brooklyn, NY")
        profile = self._profile(location="New York, NY")
        score, breakdown = score_match(person, profile)
        assert "state_match" in breakdown

    def test_verified_adds_points(self):
        person = Person("John", "Doe")
        profile_unverified = self._profile(verified=False)
        profile_verified   = self._profile(verified=True)
        score_u, _ = score_match(person, profile_unverified)
        score_v, _ = score_match(person, profile_verified)
        assert score_v > score_u

    def test_wrong_name_scores_low(self):
        person = Person("Alice", "Johnson", address="Seattle, WA", employer="Boeing")
        profile = self._profile(
            display_name="Robert Martinez",
            bio="Retired teacher",
            location="Miami, FL",
        )
        score, _ = score_match(person, profile)
        assert score < 45

    def test_score_capped_at_100(self):
        person = Person("John", "Doe", address="New York, NY", employer="Acme Corp")
        profile = self._profile(verified=True)
        score, _ = score_match(person, profile)
        assert score <= 100

    def test_confidence_thresholds(self):
        person = Person("John", "Doe", address="New York, NY", employer="Acme Corp")

        high_profile = self._profile(verified=True)
        score_h, _ = score_match(person, high_profile)

        low_profile = self._profile(
            display_name="Jonathan D.",
            bio="Just a person",
            location="",
            verified=False,
            created_year=2023,
        )
        score_l, _ = score_match(person, low_profile)

        assert score_h > score_l


# ---------------------------------------------------------------------------
# URL parsing
# ---------------------------------------------------------------------------

class TestParseTwitterUrl:
    def test_twitter_com(self):
        assert _parse_twitter_url("https://twitter.com/johndoe") == "@johndoe"

    def test_x_com(self):
        assert _parse_twitter_url("https://x.com/janedoe") == "@janedoe"

    def test_trailing_slash(self):
        assert _parse_twitter_url("https://twitter.com/user123/") == "@user123"

    def test_with_path(self):
        assert _parse_twitter_url("https://twitter.com/user/status/12345") == "@user"

    def test_reserved_word_excluded(self):
        assert _parse_twitter_url("https://twitter.com/search?q=something") is None
        assert _parse_twitter_url("https://twitter.com/intent/follow") is None

    def test_non_twitter_url(self):
        assert _parse_twitter_url("https://linkedin.com/in/johndoe") is None

    def test_empty_string(self):
        assert _parse_twitter_url("") is None


# ---------------------------------------------------------------------------
# Query builder
# ---------------------------------------------------------------------------

class TestBuildQueries:
    def test_returns_list(self):
        p = Person("John", "Smith", address="Austin, TX", employer="Dell")
        queries = _build_queries(p)
        assert isinstance(queries, list)
        assert len(queries) >= 3

    def test_most_specific_first(self):
        p = Person("John", "Smith", address="Austin, TX", employer="Dell")
        queries = _build_queries(p)
        # First query should be most specific (has both employer and city)
        assert "Dell" in queries[0] or "Austin" in queries[0]

    def test_no_employer_still_generates_queries(self):
        p = Person("John", "Smith", address="Austin, TX")
        queries = _build_queries(p)
        assert len(queries) >= 2

    def test_name_always_present(self):
        p = Person("Alice", "Walker")
        for q in _build_queries(p):
            assert "Alice" in q or "Walker" in q


# ---------------------------------------------------------------------------
# Profile builder
# ---------------------------------------------------------------------------

class TestBuildProfileFromCandidate:
    def test_builds_from_raw(self):
        cand = {
            "handle": "@testuser",
            "url": "https://twitter.com/testuser",
            "raw": {
                "display_name": "Test User",
                "bio": "Software engineer",
                "location": "San Francisco",
                "followers": 1200,
                "verified": True,
                "created_year": 2016,
                "profile_url": "https://twitter.com/testuser",
            }
        }
        profile = build_profile_from_candidate(cand)
        assert profile.handle == "@testuser"
        assert profile.display_name == "Test User"
        assert profile.followers == 1200
        assert profile.verified is True

    def test_builds_from_snippet(self):
        cand = {
            "handle": "@snippetuser",
            "title": "Snippet User (@snippetuser) · Twitter",
            "snippet": "Engineer at BigCo. Based in Chicago.",
            "url": "https://twitter.com/snippetuser",
        }
        profile = build_profile_from_candidate(cand)
        assert profile.handle == "@snippetuser"
        assert "Engineer" in profile.bio

    def test_zero_followers_default(self):
        cand = {"handle": "@x", "title": "X", "snippet": "", "url": ""}
        profile = build_profile_from_candidate(cand)
        assert profile.followers == 0


# ---------------------------------------------------------------------------
# Spreadsheet I/O
# ---------------------------------------------------------------------------

class TestLoadPersons:
    def _make_excel(self, data: dict) -> io.BytesIO:
        df = pd.DataFrame(data)
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        buf.seek(0)
        buf.name = "test.xlsx"
        return buf

    def test_loads_basic_columns(self):
        buf = self._make_excel({
            "First Name": ["Alice", "Bob"],
            "Last Name": ["Smith", "Jones"],
            "Address": ["New York, NY", "Chicago, IL"],
            "Employer": ["Acme", "BigCo"],
        })
        df, persons, warnings = load_persons(buf)
        assert len(persons) == 2
        assert persons[0].first_name == "Alice"
        assert persons[1].employer == "BigCo"

    def test_case_insensitive_columns(self):
        buf = self._make_excel({
            "FIRST NAME": ["Alice"],
            "LAST NAME": ["Smith"],
        })
        df, persons, warnings = load_persons(buf)
        assert persons[0].first_name == "Alice"

    def test_missing_address_warns(self):
        buf = self._make_excel({
            "First Name": ["Alice"],
            "Last Name": ["Smith"],
        })
        df, persons, warnings = load_persons(buf)
        assert any("address" in w.lower() for w in warnings)

    def test_missing_first_name_raises(self):
        buf = self._make_excel({"Last Name": ["Smith"]})
        with pytest.raises(ValueError, match="First Name"):
            load_persons(buf)

    def test_missing_last_name_raises(self):
        buf = self._make_excel({"First Name": ["Alice"]})
        with pytest.raises(ValueError, match="Last Name"):
            load_persons(buf)


class TestWriteResults:
    def _make_file_and_results(self):
        df = pd.DataFrame({
            "First Name": ["Alice"],
            "Last Name": ["Smith"],
        })
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        buf.seek(0)
        buf.name = "test.xlsx"

        person = Person("Alice", "Smith")
        profile = TwitterProfile(
            handle="@alicesmith",
            display_name="Alice Smith",
            bio="NYC based journalist",
            location="New York",
            followers=2300,
            verified=False,
            created_year=2015,
            profile_url="https://twitter.com/alicesmith",
        )
        result = MatchResult(
            person=person,
            profile=profile,
            score=82,
            confidence="HIGH",
            score_breakdown={"name_exact": 40, "city_match": 20, "employer_match": 25},
        )
        return buf, [result]

    def test_output_is_valid_excel(self):
        file_buf, results = self._make_file_and_results()
        out = write_results(file_buf, results)
        assert isinstance(out, io.BytesIO)
        # Verify it loads as valid Excel
        out.seek(0)
        wb_check = pd.read_excel(out, sheet_name=None)
        assert len(wb_check) >= 1

    def test_output_has_summary_sheet(self):
        file_buf, results = self._make_file_and_results()
        out = write_results(file_buf, results)
        out.seek(0)
        sheets = pd.read_excel(out, sheet_name=None)
        assert "Summary" in sheets

    def test_output_contains_handle(self):
        file_buf, results = self._make_file_and_results()
        out = write_results(file_buf, results)
        out.seek(0)
        df = pd.read_excel(out)
        assert "Twitter Handle" in df.columns
        assert df["Twitter Handle"].iloc[0] == "@alicesmith"

    def test_no_match_result(self):
        file_buf, _ = self._make_file_and_results()
        person = Person("Unknown", "Person")
        result = MatchResult(
            person=person,
            profile=None,
            score=0,
            confidence="NO MATCH",
        )
        out = write_results(file_buf, [result])
        out.seek(0)
        df = pd.read_excel(out)
        assert "✗" in str(df["Confidence"].iloc[0]) or "NO MATCH" in str(df["Confidence"].iloc[0])


# ---------------------------------------------------------------------------
# Normalisation utility
# ---------------------------------------------------------------------------

class TestNormalise:
    def test_lowercases(self):
        assert _normalise("HELLO") == "hello"

    def test_strips_special_chars(self):
        result = _normalise("Hello, World! @#$")
        assert "," not in result
        assert "@" not in result

    def test_preserves_spaces(self):
        assert "hello" in _normalise("Hello World")
        assert "world" in _normalise("Hello World")


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
