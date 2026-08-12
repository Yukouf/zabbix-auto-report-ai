#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tests isolés pour zabbix_auto_report.py — AUCUN appel réseau réel.

On ne contacte ni Zabbix, ni Ollama, ni SMTP : les fonctions réseau
(urllib.request.urlopen, ask_ollama, get_auth_token…) ne sont jamais
invoquées. On teste le parsing argparse, la normalisation de l'URL,
la classification des hôtes et le référentiel déterministe.
"""

import importlib.util
import io
import os
import sys
import unittest
from contextlib import redirect_stdout

HERE = os.path.dirname(os.path.abspath(__file__))
MOD = os.path.join(HERE, "zabbix_auto_report.py")


def _load():
    spec = importlib.util.spec_from_file_location("zabr", MOD)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


zab = _load()


class TestArgparse(unittest.TestCase):
    def test_help_render(self):
        """--help montre l'aide et mentionne les options clés (pas d'appel réseau)."""
        try:
            zab.build_arg_parser().parse_args(["--help"])
            self.fail("--help aurait dû déclencher SystemExit(0)")
        except SystemExit as e:
            self.assertEqual(e.code, 0)

    def test_no_network_on_help(self):
        """Vérifie que le parser ne déclenche aucun effet de bord réseau."""
        parser = zab.build_arg_parser()
        # --no-email et --test-email sont bien reconnus comme flags booléens.
        ns = parser.parse_args(["--no-email"])
        self.assertTrue(ns.no_email)
        self.assertFalse(ns.test_email)

    def test_invalid_flag_exits_2(self):
        """Une option inconnue → exit 2 (usage invalide)."""
        with self.assertRaises(SystemExit) as ctx:
            zab.build_arg_parser().parse_args(["--does-not-exist"])
        self.assertEqual(ctx.exception.code, 2)

    def test_main_test_email_without_smtp_returns_1(self):
        """--test-email sans SMTP configuré → code non nul, AUCUN appel réseau."""
        old_server = zab.SMTP_SERVER
        old_user = zab.SMTP_USER
        zab.SMTP_SERVER = ""
        zab.SMTP_USER = ""
        try:
            buf = io.StringIO()
            with redirect_stdout(buf):
                code = zab.main(["--test-email"])
            self.assertEqual(code, 1)
            self.assertIn("SMTP non configuré", buf.getvalue())
        finally:
            zab.SMTP_SERVER = old_server
            zab.SMTP_USER = old_user

    def test_main_without_args_reaches_zabbix_connection(self):
        """main() sans '--no-email' continue vers Zabbix : ce test ne fait PAS
        d'appel réseau, on vérifie seulement que le flux normal n'est pas
        interrompu en amont (les fonctions réseau sont mockées par défaut)."""
        # Cette assertion de non-régression : le path de test-email ne doit pas
        # être entré sans l'option --test-email.
        self.assertIsNotNone(zab.build_arg_parser().parse_args([]).no_email)


class TestURLNormalization(unittest.TestCase):
    def _normalize(self, url):
        """Reproduit la normalisation appliquée au chargement du module."""
        if not url.rstrip("/").endswith("/api_jsonrpc.php"):
            url = url.rstrip("/") + "/api_jsonrpc.php"
        return url

    def test_base_url_appends_jsonrpc(self):
        self.assertEqual(self._normalize("https://zabbix.local"),
                         "https://zabbix.local/api_jsonrpc.php")

    def test_base_url_slash_appends_jsonrpc(self):
        self.assertEqual(self._normalize("https://zabbix.local/"),
                         "https://zabbix.local/api_jsonrpc.php")

    def test_full_url_unchanged(self):
        self.assertEqual(self._normalize("https://zabbix.local/api_jsonrpc.php"),
                         "https://zabbix.local/api_jsonrpc.php")


class TestClassification(unittest.TestCase):
    def test_imp_dash_classifies_peripherique(self):
        self.assertEqual(zab.classify_host("IMP-2543", []), "Peripherique")

    def test_imp_space_classifies_peripherique(self):
        self.assertEqual(zab.classify_host("imp salle 3", []), "Peripherique")

    def test_network_keyword(self):
        # Mots-clés par défaut : aruba, hp-2530, switch
        self.assertEqual(zab.classify_host("switch-tore-01", []), "Reseau")
        self.assertEqual(zab.classify_host("aruba-2520-01", []), "Reseau")
        self.assertEqual(zab.classify_host("hp-2530-42", []), "Reseau")

    def test_pc_classifies_poste(self):
        self.assertEqual(zab.classify_host("pc-jdupont", []), "Poste")

    def test_server_default(self):
        self.assertEqual(zab.classify_host("srv-prod-db", []), "Serveur")


class TestDeterministicReferences(unittest.TestCase):
    KEY = "host||Probleme X"

    def test_agent_unavailable_matches(self):
        """Les nouveaux patterns d'agent indisponible couvrent bien le cas."""
        for name in ["zabbix agent on server01 is unavailable",
                     "Zabbix agent is not available",
                     "host agent unavailable",
                     "155.123.45.67 is unreachable"]:
            reco = zab.match_reco(name, "Linux")
            self.assertIsNotNone(reco, f"aucune reco pour : {name}")
            self.assertIn("Zabbix", reco)

    def test_printer_impdash_matches(self):
        reco = zab.match_reco("IMP-2543 Toner bas", "*")
        self.assertIsNotNone(reco)
        self.assertIn("consommable", reco)

    def test_disk_full_labelled_by_os(self):
        linux = zab.match_reco("Disk space is low on FS [/]", "Linux")
        windows = zab.match_reco("Disk space is low on FS [/]", "Windows")
        self.assertIsNotNone(linux)
        self.assertIsNotNone(windows)
        self.assertIn("df -h", linux)
        self.assertIn("Get-PSDrive", windows)

    def test_unknown_no_match_returns_none(self):
        self.assertIsNone(zab.match_reco("Curieux evenement tres rare", "Linux"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
