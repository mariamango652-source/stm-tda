import React, { useState, useEffect, useMemo, useCallback } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

const SB = createClient("https://ssqrpfrbxzafcykyusmp.supabase.co", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNzcXJwZnJieHphZmN5a3l1c21wIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MTA4NzE1OCwiZXhwIjoyMDk2NjYzMTU4fQ.lrTKUXGBrb4jGO8gOIX0m6AHNG2QGidC5fryiTdqErw");

/* ═══ DATA ═══ */
const CL="Л-1|Л-1А|Л-2А|Л-2Б|Л-3|Л-3А|Л-3Б|Л-4|Л-4А|Л-5|Л-5А|Л-5Б|Л-6|Л-6А|Л-7|Л-8|Л-9|Л-9А|Л-10|Л-10А|Л-11|Л-11А|Л-11А + 30%|Л-12|Л-12А|Л-12Б|Л-13|Л-13А|Л-14|Л-15|Л-15А|Л-15Б|Л-16|Л-17|Л-18|Л-18А|Л-19|Л-19А|Л-19Б|Л-20|Л-21";
const CC="Ceramic 70|Ceramic 65|Amber 75|Mariil 110|Palazzo 230|Caramel 65|Calvados 6|Sahara 4|Magma 80|Cameo 75|Bordeaux 10|Mariil 80|Granit 5|Granit 10|Granit 15|Granit 20|Granit 25|Granit 30|Granit 35|Granit 40|Granit 45|Granit 50|Granit 55|Granit 60|Ferro 5|Ferro 10|Ferro 15|Ferro 20|Ferro 25|Ferro 30|Ferro 35|Ferro 40|Ferro 45|Ferro 50|Ferro 55|Ferro 60|Jura 5|Jura 10|Jura 15|Jura 20|Jura 25|Jura 30|Jura 35|Jura 40|Jura 45|Jura 50|Jura 55|Jura 60|Venato 5|Venato 10|Venato 15|Venato 20|Venato 25|Venato 30|Venato 35|Venato 40|Venato 45|Venato 50|Venato 55|Venato 60|Rose 5|Rose 10|Rose 15|Rose 20|Rose 25|Rose 30|Rose 35|Rose 40|Rose 45|Rose 50|Rose 55|Rose 60|Rose 65|Rose 70|Rose 75|Rose 80|Rose 85|Rose 90|Rose 95|Rose 100|Rose 105|Rose 110|Rose 115|Rose 120|Baccara 5|Baccara 10|Baccara 15|Baccara 20|Baccara 25|Baccara 30|Barolo 5|Barolo 10|Barolo 15|Barolo 20|Barolo 25|Barolo 30|Barolo 35|Barolo 40|Barolo 45|Barolo 50|Barolo 55|Barolo 60|Barolo 65|Barolo 70|Barolo 75|Barolo 80|Barolo 85|Barolo 90|Barolo 95|Barolo 100|Barolo 105|Barolo 110|Barolo 115|Barolo 120|Barolo 125|Barolo 130|Barolo 135|Barolo 140|Barolo 145|Barolo 150|Venezia 5|Venezia 10|Venezia 15|Venezia 20|Venezia 25|Venezia 30|Grenadin 5|Grenadin 10|Grenadin 15|Grenadin 20|Grenadin 25|Grenadin 30|Grenadin 35|Grenadin 40|Grenadin 45|Grenadin 50|Grenadin 55|Grenadin 60|Grenadin 65|Grenadin 70|Grenadin 75|Grenadin 80|Grenadin 85|Grenadin 90|Grenadin 95|Grenadin 100|Grenadin 105|Grenadin 110|Grenadin 115|Grenadin 120|Grenadin 125|Grenadin 130|Grenadin 135|Grenadin 140|Grenadin 145|Grenadin 150|Bordeaux 5|Bordeaux 15|Bordeaux 20|Bordeaux 25|Bordeaux 30|Bordeaux 35|Bordeaux 40|Bordeaux 45|Bordeaux 50|Bordeaux 55|Bordeaux 60|Bordeaux 65|Bordeaux 70|Bordeaux 75|Bordeaux 80|Bordeaux 85|Bordeaux 90|Magma 5|Magma 10|Magma 15|Magma 20|Magma 25|Magma 30|Magma 35|Magma 40|Magma 45|Magma 50|Magma 55|Magma 60|Magma 65|Magma 70|Magma 75|Magma 85|Magma 90|Magma 95|Magma 100|Magma 105|Magma 110|Magma 115|Magma 120|Magma 125|Magma 130|Magma 135|Magma 140|Magma 145|Magma 150|Cameo 5|Cameo 10|Cameo 15|Cameo 20|Cameo 25|Cameo 30|Cameo 35|Cameo 40|Cameo 45|Cameo 50|Cameo 55|Cameo 60|Cameo 65|Cameo 70|Cameo 80|Cameo 85|Cameo 90|Cameo 95|Cameo 100|Cameo 105|Cameo 110|Cameo 115|Cameo 120|Cameo 125|Cameo 130|Cameo 135|Cameo 140|Cameo 145|Cameo 150|Cameo 155|Cameo 160|Cameo 165|Cameo 170|Cameo 175|Cameo 180|Lachs 5|Lachs 10|Lachs 15|Lachs 20|Lachs 25|Lachs 30|Lachs 35|Lachs 40|Lachs 45|Lachs 50|Lachs 55|Lachs 60|Lachs 65|Lachs 70|Lachs 75|Lachs 80|Lachs 85|Lachs 90|Lachs 95|Lachs 100|Lachs 105|Lachs 110|Lachs 115|Lachs 120|Aprico 5|Aprico 10|Aprico 15|Aprico 20|Aprico 25|Aprico 30|Aprico 35|Aprico 40|Aprico 45|Aprico 50|Aprico 55|Aprico 60|Aprico 65|Aprico 70|Aprico 75|Aprico 80|Aprico 85|Aprico 90|Aprico 95|Aprico 100|Aprico 105|Aprico 110|Aprico 115|Aprico 120|Aprico 125|Aprico 130|Aprico 135|Aprico 140|Aprico 145|Aprico 150|Aprico 155|Aprico 160|Aprico 165|Aprico 170|Aprico 175|Aprico 180|Papaya 5|Papaya 10|Papaya 15|Papaya 20|Papaya 25|Papaya 30|Papaya 35|Papaya 40|Papaya 45|Papaya 50|Papaya 55|Papaya 60|Papaya 65|Papaya 70|Papaya 75|Papaya 80|Papaya 85|Papaya 90|Papaya 95|Papaya 100|Papaya 105|Papaya 110|Papaya 115|Papaya 120|Papaya 125|Papaya 130|Papaya 135|Papaya 140|Papaya 145|Papaya 150|Ceramic 5|Ceramic 10|Ceramic 15|Ceramic 20|Ceramic 25|Ceramic 30|Ceramic 35|Ceramic 40|Ceramic 45|Ceramic 50|Ceramic 55|Ceramic 60|Ceramic 75|Ceramic 80|Ceramic 85|Ceramic 90|Amber 5|Amber 10|Amber 15|Amber 20|Amber 25|Amber 30|Amber 35|Amber 40|Amber 45|Amber 50|Amber 55|Amber 60|Amber 65|Amber 70|Amber 80|Amber 85|Amber 90|Amber 95|Amber 100|Amber 105|Amber 110|Amber 115|Amber 120|Amber 125|Amber 130|Amber 135|Amber 140|Amber 145|Amber 150|Siena 5|Siena 10|Siena 15|Siena 20|Siena 25|Siena 30|Siena 35|Siena 40|Siena 45|Siena 50|Siena 55|Siena 60|Siena 65|Siena 70|Siena 75|Siena 80|Siena 85|Siena 90|Siena 95|Siena 100|Siena 105|Siena 110|Siena 115|Siena 120|Siena 125|Siena 130|Siena 135|Siena 140|Siena 145|Siena 150|Siena 155|Siena 160|Siena 165|Siena 170|Siena 175|Siena 180|Onyx 5|Onyx 10|Onyx 15|Onyx 20|Onyx 25|Onyx 30|Onyx 35|Onyx 40|Onyx 45|Onyx 50|Onyx 55|Onyx 60|Onyx 65|Onyx 70|Onyx 75|Onyx 80|Onyx 85|Onyx 90|Onyx 95|Onyx 100|Onyx 105|Onyx 110|Onyx 115|Onyx 120|Onyx 125|Onyx 130|Onyx 135|Onyx 140|Onyx 145|Onyx 150|Onyx 155|Onyx 160|Onyx 165|Onyx 170|Onyx 175|Onyx 180|Onyx 185|Onyx 190|Onyx 195|Onyx 200|Onyx 205|Onyx 210|Onyx 215|Onyx 220|Onyx 225|Onyx 230|Onyx 235|Onyx 240|Marill 5|Marill 10|Marill 15|Marill 20|Marill 25|Marill 30|Marill 35|Marill 40|Marill 45|Marill 50|Marill 55|Marill 60|Marill 65|Marill 70|Marill 75|Marill 80|Marill 85|Marill 90|Marill 95|Marill 100|Marill 105|Marill 110|Marill 115|Marill 120|Marill 125|Marill 130|Marill 135|Marill 140|Marill 145|Marill 150|Palazzo 5|Palazzo 10|Palazzo 15|Palazzo 20|Palazzo 25|Palazzo 30|Palazzo 35|Palazzo 40|Palazzo 45|Palazzo 50|Palazzo 55|Palazzo 60|Palazzo 65|Palazzo 70|Palazzo 75|Palazzo 80|Palazzo 85|Palazzo 90|Palazzo 95|Palazzo 100|Palazzo 105|Palazzo 110|Palazzo 115|Palazzo 120|Palazzo 125|Palazzo 130|Palazzo 135|Palazzo 140|Palazzo 145|Palazzo 150|Palazzo 155|Palazzo 160|Palazzo 165|Palazzo 170|Palazzo 175|Palazzo 180|Palazzo 185|Palazzo 190|Palazzo 195|Palazzo 200|Palazzo 205|Palazzo 210|Palazzo 215|Palazzo 220|Palazzo 225|Palazzo 235|Palazzo 240|Palazzo 245|Palazzo 250|Palazzo 255|Palazzo 260|Palazzo 265|Palazzo 270|Palazzo 275|Palazzo 280|Palazzo 285|Palazzo 290|Palazzo 295|Palazzo 300|Palazzo 305|Palazzo 310|Palazzo 315|Palazzo 320|Palazzo 325|Palazzo 330|Palazzo 335|Palazzo 340|Palazzo 345|Palazzo 350|Palazzo 355|Palazzo 360|Curcuma 5|Curcuma 10|Curcuma 15|Curcuma 20|Curcuma 25|Curcuma 30|Curcuma 35|Curcuma 40|Curcuma 45|Curcuma 50|Curcuma 55|Curcuma 60|Curcuma 65|Curcuma 70|Curcuma 75|Curcuma 80|Curcuma 85|Curcuma 90|Curcuma 95|Curcuma 100|Curcuma 105|Curcuma 110|Curcuma 115|Curcuma 120|Ginster 5|Ginster 10|Ginster 15|Ginster 20|Ginster 25|Ginster 30|Ginster 35|Ginster 40|Ginster 45|Ginster 50|Ginster 55|Ginster 60|Ginster 65|Ginster 70|Ginster 75|Ginster 80|Ginster 85|Ginster 90|Ginster 95|Ginster 100|Ginster 105|Ginster 110|Ginster 115|Ginster 120|Ginster 125|Ginster 130|Ginster 135|Ginster 140|Ginster 145|Ginster 150|Curry 5|Curry 10|Curry 15|Curry 20|Curry 25|Curry 30|Curry 35|Curry 40|Curry 45|Curry 50|Curry 55|Curry 60|Curry 65|Curry 70|Curry 75|Curry 80|Curry 85|Curry 90|Curry 95|Curry 100|Curry 105|Curry 110|Curry 115|Curry 120|Curry 125|Curry 130|Curry 135|Curry 140|Curry 145|Curry 150|Citrus 5|Citrus 10|Citrus 15|Citrus 20|Citrus 25|Citrus 30|Citrus 35|Citrus 40|Citrus 45|Citrus 50|Citrus 55|Citrus 60|Citrus 65|Citrus 70|Citrus 75|Citrus 80|Citrus 85|Citrus 90|Citrus 95|Citrus 100|Citrus 105|Citrus 110|Citrus 115|Citrus 120|Melisse 5|Melisse 10|Melisse 15|Melisse 20|Melisse 25|Melisse 30|Melisse 35|Melisse 40|Melisse 45|Melisse 50|Melisse 55|Melisse 60|Melisse 65|Melisse 70|Melisse 75|Melisse 80|Melisse 85|Melisse 90|Melisse 95|Melisse 100|Melisse 105|Melisse 110|Melisse 115|Melisse 120|Soja 5|Soja 10|Soja 15|Soja 20|Soja 25|Soja 30|Tundra 5|Tundra 10|Tundra 15|Tundra 20|Tundra 25|Tundra 30|Tundra 35|Tundra 40|Tundra 45|Tundra 50|Tundra 55|Tundra 60|Tundra 65|Tundra 70|Tundra 75|Tundra 80|Tundra 85|Tundra 90|Tundra 95|Tundra 100|Tundra 105|Tundra 110|Tundra 115|Tundra 120|Tundra 125|Tundra 130|Tundra 135|Tundra 140|Tundra 145|Tundra 150|Mai 5|Mai 10|Mai 15|Mai 20|Mai 25|Mai 30|Mai 35|Mai 40|Mai 45|Mai 50|Mai 55|Mai 60|Mai 65|Mai 70|Mai 75|Mai 80|Mai 85|Mai 90|Jade 5|Jade 10|Jade 15|Jade 20|Jade 25|Jade 30|Jade 35|Jade 40|Jade 45|Jade 50|Jade 55|Jade 60|Jade 65|Jade 70|Jade 75|Jade 80|Jade 85|Jade 90|Jade 95|Jade 100|Jade 105|Jade 110|Jade 115|Jade 120|Pinie 5|Pinie 10|Pinie 15|Pinie 20|Pinie 25|Pinie 30|Pinie 35|Pinie 40|Pinie 45|Pinie 50|Pinie 55|Pinie 60|Pinie 65|Pinie 70|Pinie 75|Pinie 80|Pinie 85|Pinie 90|Pinie 95|Pinie 100|Pinie 105|Pinie 110|Pinie 115|Pinie 120|Moos 5|Moos 10|Moos 15|Moos 20|Moos 25|Moos 30|Moos 35|Moos 40|Moos 45|Moos 50|Moos 55|Moos 60|Moos 65|Moos 70|Moos 75|Moos 80|Moos 85|Moos 90|Oase 5|Oase 10|Oase 15|Oase 20|Oase 25|Oase 30|Oase 35|Oase 40|Oase 45|Oase 50|Oase 55|Oase 60|Oase 65|Oase 70|Oase 75|Oase 80|Oase 85|Oase 90|Oase 95|Oase 100|Oase 105|Oase 110|Oase 115|Oase 120|Oase 125|Oase 130|Oase 135|Oase 140|Oase 145|Oase 150|Agave 5|Agave 10|Agave 15|Agave 20|Agave 25|Agave 30|Agave 35|Agave 40|Agave 45|Agave 50|Agave 55|Agave 60|Agave 65|Agave 70|Agave 75|Agave 80|Agave 85|Agave 90|Agave 95|Agave 100|Agave 105|Agave 110|Agave 115|Agave 120|Agave 125|Agave 130|Agave 135|Agave 140|Agave 145|Agave 150|Malachit 5|Malachit 10|Malachit 15|Malachit 20|Malachit 25|Malachit 30|Malachit 35|Malachit 40|Malachit 45|Malachit 50|Malachit 55|Malachit 60|Malachit 65|Malachit 70|Malachit 75|Malachit 80|Malachit 85|Malachit 90|Malachit 95|Malachit 100|Malachit 105|Malachit 110|Malachit 115|Malachit 120|Mint 5|Mint 10|Mint 15|Mint 20|Mint 25|Mint 30|Mint 35|Mint 40|Mint 45|Mint 50|Mint 55|Mint 60|Patina 5|Patina 10|Patina 15|Patina 20|Patina 25|Patina 30|Patina 35|Patina 40|Patina 45|Patina 50|Patina 55|Patina 60|Patina 65|Patina 70|Patina 75|Patina 80|Patina 85|Patina 90|Patina 95|Patina 100|Patina 105|Patina 110|Patina 115|Patina 120|Verona 5|Verona 10|Verona 15|Verona 20|Verona 25|Verona 30|Verona 35|Verona 40|Verona 45|Verona 50|Verona 55|Verona 60|Verona 65|Verona 70|Verona 75|Verona 80|Verona 85|Verona 90|Verona 95|Verona 100|Verona 105|Verona 110|Verona 115|Verona 120|Verona 125|Verona 130|Verona 135|Verona 140|Verona 145|Verona 150|Arctis 5|Arctis 10|Arctis 15|Arctis 20|Arctis 25|Arctis 30|Arctis 35|Arctis 40|Arctis 45|Arctis 50|Arctis 55|Arctis 60|Arctis 65|Arctis 70|Arctis 75|Arctis 80|Arctis 85|Arctis 90|Arctis 95|Arctis 100|Arctis 105|Arctis 110|Arctis 115|Arctis 120|Coelin 5|Coelin 10|Coelin 15|Coelin 20|Coelin 25|Coelin 30|Coelin 35|Coelin 40|Coelin 45|Coelin 50|Coelin 55|Coelin 60|Coelin 65|Coelin 70|Coelin 75|Coelin 80|Coelin 85|Coelin 90|Coelin 95|Coelin 100|Coelin 105|Coelin 110|Coelin 115|Coelin 120|Pacific 5|Pacific 10|Pacific 15|Pacific 20|Pacific 25|Pacific 30|Pacific 35|Pacific 40|Pacific 45|Pacific 50|Pacific 55|Pacific 60|Pacific 65|Pacific 70|Pacific 75|Pacific 80|Pacific 85|Pacific 90|Pacific 95|Pacific 100|Pacific 105|Pacific 110|Pacific 115|Pacific 120|Pacific 125|Pacific 130|Pacific 135|Pacific 140|Pacific 145|Pacific 150|Pacific 155|Pacific 160|Pacific 165|Pacific 170|Pacific 175|Pacific 180|Lago 5|Lago 10|Lago 15|Lago 20|Lago 25|Lago 30|Lago 35|Lago 40|Lago 45|Lago 50|Lago 55|Lago 60|Lago 65|Lago 70|Lago 75|Lago 80|Lago 85|Lago 90|Lago 95|Lago 100|Lago 105|Lago 110|Lago 115|Lago 120|Lazur 5|Lazur 10|Lazur 15|Lazur 20|Lazur 25|Lazur 30|Lazur 35|Lazur 40|Lazur 45|Lazur 50|Lazur 55|Lazur 60|Lazur 65|Lazur 70|Lazur 75|Lazur 80|Lazur 85|Lazur 90|Lazur 95|Lazur 100|Lazur 105|Lazur 110|Lazur 115|Lazur 120|Lazur 125|Lazur 130|Lazur 135|Lazur 140|Lazur 145|Lazur 150|Lazur 155|Lazur 160|Lazur 165|Lazur 170|Lazur 175|Lazur 180|Lazur 185|Lazur 190|Lazur 195|Lazur 200|Lazur 205|Lazur 210|Saphir 5|Saphir 10|Saphir 15|Saphir 20|Saphir 25|Saphir 30|Saphir 35|Saphir 40|Saphir 45|Saphir 50|Saphir 55|Saphir 60|Saphir 65|Saphir 70|Saphir 75|Saphir 80|Saphir 85|Saphir 90|Saphir 95|Saphir 100|Saphir 105|Saphir 110|Saphir 115|Saphir 120|Aquarell 5|Aquarell 10|Aquarell 15|Aquarell 20|Aquarell 25|Aquarell 30|Aquarell 35|Aquarell 40|Aquarell 45|Aquarell 50|Aquarell 55|Aquarell 60|Aquarell 65|Aquarell 70|Aquarell 75|Aquarell 80|Aquarell 85|Aquarell 90|Aquarell 95|Aquarell 100|Aquarell 105|Aquarell 110|Aquarell 115|Aquarell 120|Lavendel 5|Lavendel 10|Lavendel 15|Lavendel 20|Lavendel 25|Lavendel 30|Lavendel 35|Lavendel 40|Lavendel 45|Lavendel 50|Lavendel 55|Lavendel 60|Lavendel 65|Lavendel 70|Lavendel 75|Lavendel 80|Lavendel 85|Lavendel 90|Lavendel 95|Lavendel 100|Lavendel 105|Lavendel 110|Lavendel 115|Lavendel 120|Lavendel 125|Lavendel 130|Lavendel 135|Lavendel 140|Lavendel 145|Lavendel 150|Lavendel 155|Lavendel 160|Lavendel 165|Lavendel 170|Lavendel 175|Lavendel 180|Velvet 5|Velvet 10|Velvet 15|Velvet 20|Velvet 25|Velvet 30|Velvet 35|Velvet 40|Velvet 45|Velvet 50|Velvet 55|Velvet 60|Velvet 65|Velvet 70|Velvet 75|Velvet 80|Velvet 85|Velvet 90|Velvet 95|Velvet 100|Velvet 105|Velvet 110|Velvet 115|Velvet 120|Viola 5|Viola 10|Viola 15|Viola 20|Viola 25|Viola 30|Viola 35|Viola 40|Viola 45|Viola 50|Viola 55|Viola 60|Viola 65|Viola 70|Viola 75|Viola 80|Viola 85|Viola 90|Viola 95|Viola 100|Viola 105|Viola 110|Viola 115|Viola 120|Flamenco 5|Flamenco 10|Flamenco 15|Flamenco 20|Flamenco 25|Flamenco 30|Flamenco 35|Flamenco 40|Flamenco 45|Flamenco 50|Flamenco 55|Flamenco 60|Flamenco 65|Flamenco 70|Flamenco 75|Flamenco 80|Flamenco 85|Flamenco 90|Flamenco 95|Flamenco 100|Flamenco 105|Flamenco 110|Flamenco 115|Flamenco 120|Flamenco 125|Flamenco 130|Flamenco 135|Flamenco 140|Flamenco 145|Flamenco 150|Diva 5|Diva 10|Diva 15|Diva 20|Diva 25|Diva 30|Blu 5|Blu 10|Blu 15|Blu 20|Blu 25|Blu 30|Peru 5|Peru 10|Peru 15|Peru 20|Peru 25|Peru 30|Verdo 5|Verdo 10|Verdo 15|Verdo 20|Verdo 25|Verdo 30|Oliv 5|Oliv 10|Oliv 15|Oliv 20|Oliv 25|Oliv 30|Brasil 5|Brasil 10|Brasil 15|Brasil 20|Brasil 25|Brasil 30|Rubin 5|Rubin 10|Rubin 15|Rubin 20|Rubin 25|Rubin 30|Pink 5|Pink 10|Pink 15|Pink 20|Pink 25|Pink 30|Laser 5|Laser 10|Laser 15|Laser 20|Laser 25|Laser 30|Laser 35|Laser 40|Laser 45|Laser 50|Laser 55|Laser 60|Laser 65|Laser 70|Laser 75|Laser 80|Laser 85|Laser 90|Grau 5|Grau 10|Grau 15|Grau 20|Grau 25|Grau 30|Grau 35|Grau 40|Grau 45|Grau 50|Grau 55|Grau 60|Schwarz-Grau|Sepia-Grau|Umbra-Grau|Graphit-Grau|Bläulich-Grau|Schiefer-Grau|Off White 5|Off White 10|Off White 35|Off White 45|Off White 55|Off White 60|Quarz 10|Quarz 15|Quarz 20|Quarz 25|Quarz 30|Quarz 45|Quarz 50|Quarz 55|Quarz 60|Korall 5|Korall 10|Korall 15|Korall 20|Korall 25|Korall 30|Korall 35|Korall 40|Korall 45|Korall 50|Korall 55|Korall 60|Korall 65|Korall 70|Korall 75|Korall 80|Korall 85|Korall 90|Korall 95|Korall 100|Korall 105|Korall 110|Korall 115|Korall 120|Korall 125|Korall 130|Korall 135|Korall 140|Korall 145|Korall 150|Korall 155|Korall 160|Korall 165|Korall 170|Korall 175|Korall 180|Perl 5|Perl 10|Perl 15|Perl 20|Perl 25|Perl 30|Perl 40|Perl 45|Perl 50|Perl 55|Perl 60|Ceramic 95|Ceramic 100|Ceramic 105|Kreide 5|Kreide 10|Kreide 15|Kreide 20|Kreide 25|Kreide 30|Falb 5|Falb 10|Falb 15|Falb 20|Falb 25|Falb 30|Falb 35|Falb 40|Falb 45|Falb 50|Falb 55|Falb 60|Falb 65|Falb 70|Falb 75|Falb 80|Caramel 5|Caramel 10|Caramel 15|Caramel 20|Caramel 25|Caramel 30|Caramel 35|Caramel 40|Caramel 45|Caramel 50|Caramel 55|Caramel 60|Caramel 70|Caramel 75|Caramel 80|Caramel 85|Caramel 90|Caramel 95|Caramel 100|Caramel 105|Caramel 110|Caramel 115|Caramel 120|Herba 5|Herba 10|Herba 15|Herba 20|Herba 25|Herba 30|Herba 40|Herba 45|Herba 50|Herba 55|Herba 60|Moos 95|Moos 100|Moos 105|Moos 110|Moos 115|Moos 120|Palm 5|Palm 10|Palm 15|Palm 20|Palm 25|Palm 30|Palm 35|Palm 40|Palm 45|Palm 50|Palm 55|Palm 60|Palm 65|Palm 70|Palm 75|Mint 65|Mint 70|Mint 75|Mint 80|Mint 85|Mint 90|Marin 5|Marin 10|Marin 15|Marin 20|Marin 25|Marin 30|Marin 35|Marin 40|Marin 45|Marin 50|Marin 55|Marin 60|Marin 65|Marin 70|Marin 75|Marin 80|Marin 85|Marin 90|Astra 5|Astra 10|Astra 15|Astra 20|Astra 25|Astra 30|Astra 35|Astra 40|Astra 45|Astra 50|Astra 55|Astra 60|Astra 65|Astra 70|Astra 75|Astra 80|Litho 5|Litho 10|Litho 15|Litho 20|Litho 25|Litho 30|Amethyst 5|Amethyst 10|Amethyst 15|Amethyst 20|Amethyst 25|Amethyst 30|Amethyst 35|Amethyst 40|Amethyst 45|Amethyst 50|Amethyst 55|Amethyst 60|Amethyst 65|Amethyst 70|Amethyst 75|Amethyst 80|Amethyst 85|Amethyst 90|Amethyst 95|Amethyst 100|Amethyst 105|Amethyst 110|Amethyst 115|Amethyst 120|Atmo 5|Atmo 10|Atmo 15|Atmo 20|Atmo 25|Clematis 5|Clematis 10|Clematis 15|Clematis 20|Clematis 25|Clematis 30|Clematis 35|Clematis 40|Clematis 45|Clematis 50|Clematis 55|Clematis 60|Clematis 65|Clematis 70|Clematis 75|Clematis 80|Clematis 85|Clematis 90|Alucryl 5|Alucryl 10|Alucryl 15|Alucryl 20|Alucryl 25|Alucryl 30|Alucryl 35|Alucryl 40|Alucryl 45|Alucryl 50|Alucryl 55|Alucryl 60|Alucryl 65|Alucryl 70|Alucryl 75|Alucryl 80|Alucryl 85|Alucryl 90";
const CN="NSC S 1030-Y20R|NCS S 2020-Y10R|NSC S 0500-N|NSC S 3500-N|NSC S 6000-N|NSC S 9000-N|NSC S 0502-B50G|NSC S 1002-R50B|NSC S 1502-Y50R|NSC S 1502-G50Y|NSC S 1002-R|NSC S 6502-R|NSC S 2502-B|NSC S 8502-B|NSC S 4502-G|NSC S 7005-Y20R|NSC S 3005-Y50R|NSC S 6005-Y80R|NSC S 2005-R20B|NSC S 8005-R20B|NSC S 4005-R50B|NSC S 7005-R80B|NSC S 3005-B20G|NSC S 8505-B20G|NSC S 5005-B80G|NSC S 1005-G20Y|NSC S 7005-G20Y|NSC S 3005-G50Y|NSC S 0505-G80Y|NSC S 6005-G80Y|S 4010-Y50R|S 1070-G60Y|S 4020-B70G|S 1500-N|S 4500-N|S 7000-N|S 0502-Y50R|S 0502-G50Y|S 1002-B50G|S 1502-R50B|S 1002-Y|S 2502-R|S 8502-R|S 4502-B|S 1002-G|S 6502-G";
const RS=new Set(["Л-3","Л-5","Л-5А","Л-5Б","Л-6","Л-11","Л-11А","Л-11А + 30%","Л-12","Л-12А","Л-15А","Л-21","Ceramic 70","Ceramic 65","Amber 75","Mariil 110","Palazzo 230","Caramel 65","Calvados 6","Sahara 4","Magma 80","Cameo 75","Bordeaux 10","Mariil 80","NSC S 1030-Y20R","NCS S 2020-Y10R"]);
const RA=new Set(["Л-15","Л-17"]);
function buildDB(){const db=[];const add=(p,raw)=>raw.split("|").forEach(c=>{if(c)db.push({palette:p,code:c})});add("Ленинградская",CL);add("Caparol",CC);add("NCS",CN);return db;}
const ST={new:"Новый",accepted:"Принят",waiting:"Ожидает подбора",inwork:"В работе",ready:"Готов",received:"Получен",shipped:"Отгружен",cancelled:"Отменён"};
const SC={new:"#dc2626",accepted:"#ea580c",waiting:"#d97706",inwork:"#dc2626",ready:"#16a34a",received:"#2563eb",shipped:"#71717a",cancelled:"#18181b"};
const SNEXT={new:"accepted",accepted:"inwork",waiting:"inwork",inwork:"ready",ready:"received",received:"shipped"};
const fmtDate=d=>d?new Date(d).toLocaleDateString("ru-RU",{day:"2-digit",month:"2-digit",year:"numeric"}):"—";
const fmtDateShort=d=>d?new Date(d).toLocaleDateString("ru-RU",{day:"2-digit",month:"2-digit"}):"—";
const todayStr=()=>new Date().toISOString().slice(0,10);
const LS={get:(k,d)=>{try{const v=localStorage.getItem("stm_tda_"+k);return v?JSON.parse(v):d}catch{return d}},set:(k,v)=>{try{localStorage.setItem("stm_tda_"+k,JSON.stringify(v))}catch{}}};

const BG="#f2f3f5";const W="#ffffff";const INK="#0a0a0a";const INK2="#1a1a1a";const G100="#f2f3f5";const G200="#e8e8ec";const G300="#d0d0d8";const G400="#9898a6";const G500="#6b6b7a";const RED="#ef3124";const REDL="#fff0ef";const GRN="#16a34a";const GRNL="#f0fdf4";const ORG="#ea580c";const ORGL="#fff7ed";const BLU="#2563eb";const BLUL="#eff6ff";const NAVY="#1e3a5f";const FD="'Inter',-apple-system,'Segoe UI',sans-serif";
const FF="'Inter',-apple-system,'Segoe UI',sans-serif";const R=20;const SH="0 4px 20px rgba(0,0,0,.06),0 1px 3px rgba(0,0,0,.04)";
/* Карточка — только тень, без рамки, как у Альфа-Банка */
const C={background:W,borderRadius:R,padding:20,boxShadow:SH};
/* Поле ввода — underline стиль */
const I={width:"100%",padding:"14px 0",border:"none",borderBottom:`1.5px solid ${G300}`,borderRadius:0,fontSize:16,boxSizing:"border-box",fontFamily:FF,color:INK,background:"transparent",outline:"none"};
/* Поле с рамкой (поиск и т.п.) */
const IB={width:"100%",padding:"13px 16px",border:`1.5px solid ${G200}`,borderRadius:14,fontSize:15,boxSizing:"border-box",fontFamily:FF,color:INK,background:W,outline:"none"};
/* Основная кнопка — Альфа-стиль */
const BTN={width:"100%",padding:"17px",background:RED,color:W,border:"none",borderRadius:14,fontSize:16,fontWeight:700,cursor:"pointer",fontFamily:FF,letterSpacing:0.2};

function Badge({status}){return <span style={{background:SC[status]+"18",color:SC[status],padding:"4px 12px",borderRadius:50,fontSize:11,fontWeight:700,letterSpacing:0.2}}>{ST[status]}</span>}
function TypeBadge({type}){const c=type==="образец"?RED:NAVY;const bg=type==="образец"?REDL:BLUL;return <span style={{background:bg,color:c,padding:"3px 10px",borderRadius:50,fontSize:11,fontWeight:700}}>{type==="образец"?"Образец":"Партия"}</span>}
function RB({has,label}){return <span style={{padding:"4px 12px",borderRadius:50,fontSize:11,fontWeight:700,color:has?GRN:RED,background:has?GRNL:REDL}}>{label}: {has?"есть":"нет"}</span>}
function Pill({active,onClick,children,c=RED}){return <button onClick={onClick} style={{padding:"8px 18px",borderRadius:50,border:"none",background:active?INK:G200,color:active?W:G500,fontSize:13,fontWeight:active?700:500,cursor:"pointer",whiteSpace:"nowrap",fontFamily:FF,transition:"all .15s"}}>{children}</button>}
function Tag({children}){return <div style={{fontSize:11,fontWeight:700,color:G400,marginBottom:10,letterSpacing:1.2,textTransform:"uppercase"}}>{children}</div>}
function Chevron({open}){return <span style={{color:G400,fontSize:11,transition:"transform .2s",display:"inline-block",transform:open?"rotate(180deg)":"rotate(0deg)"}}>▾</span>}

/* ═══ COLLAPSIBLE SECTION ═══ */
function Section({title,children,defaultOpen=false,accent=G400}){
  const [open,setOpen]=useState(defaultOpen);
  return <div style={{...C,marginBottom:8,padding:0,overflow:"hidden"}}>
    <div onClick={()=>setOpen(!open)} style={{padding:"12px 16px",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:open?`1px solid ${G200}`:"none"}}>
      <div style={{fontSize:12,fontWeight:800,color:accent,letterSpacing:.5,textTransform:"uppercase"}}>{title}</div>
      <Chevron open={open}/>
    </div>
    {open&&<div style={{padding:"12px 16px 14px"}}>{children}</div>}
  </div>;
}

export default function App(){
  const [colors,setColors]=useState(()=>buildDB());
  const [page,setPage]=useState("dash");
  const [orders,setOrders]=useState([]);
  const [recSync,setRecSync]=useState({});
  const [did,setDid]=useState(null);
  const [user]=useState("Мэри");
  const [loading,setLoading]=useState(true);

  // Загружаем кастомные цвета из базы сразу при монтировании
  useEffect(()=>{(async()=>{
    const{data:dbColors}=await SB.from("colors").select("palette,code,silicate_recipe,acrylic_recipe");
    if(dbColors&&dbColors.length>0){
      setColors(prev=>{
        const existing=new Set(prev.map(c=>c.palette+"|"+c.code));
        const toAdd=dbColors.filter(c=>c.palette&&c.code&&!existing.has(c.palette+"|"+c.code));
        return toAdd.length>0?[...prev,...toAdd]:prev;
      });
      setRecSync(prev=>{
        const next={...prev};
        dbColors.forEach(c=>{
          if(c.silicate_recipe)next[c.code+"__силикат"]=true;
          if(c.acrylic_recipe)next[c.code+"__акрил"]=true;
        });
        return next;
      });
    }
  })()},[]);

  useEffect(()=>{(async()=>{
    try{
      const{data:od,error:oe}=await SB.from("orders").select("*").order("created_at",{ascending:false});
      console.log("Orders load:",od?"ok "+od.length:"null",oe||"");
      const orderList=od||[];
      let cmtMap={};
      if(orderList.length>0){
        const{data:cm}=await SB.from("order_comments").select("*").in("order_id",orderList.map(o=>o.id)).order("created_at");
        (cm||[]).forEach(c=>{if(!cmtMap[c.order_id])cmtMap[c.order_id]=[];cmtMap[c.order_id].push(c)});
      }
      // Load KP attachments per order
      let kpMap={};
      try{
        const{data:kp}=await SB.from("order_kp").select("*").in("order_id",orderList.map(o=>o.id));
        (kp||[]).forEach(k=>{if(!kpMap[k.order_id])kpMap[k.order_id]=[];kpMap[k.order_id].push(k)});
      }catch(e){console.log("order_kp table may not exist yet:",e.message)}
      setOrders(orderList.map(o=>({...o,comments:cmtMap[o.id]||[],kp:kpMap[o.id]||[]})));
      // Inject palette/code from orders so colors used in past orders always appear in selector
      setColors(prev=>{
        const existing=new Set(prev.map(c=>c.palette+"|"+c.code));
        const toAdd=orderList.filter(o=>o.palette&&o.color_code&&!existing.has(o.palette+"|"+o.color_code))
          .map(o=>({palette:o.palette,code:o.color_code}));
        const deduped=toAdd.filter((x,i)=>toAdd.findIndex(y=>y.palette===x.palette&&y.code===x.code)===i);
        return deduped.length>0?[...prev,...deduped]:prev;
      });
      const{data:cd}=await SB.from("colors").select("code,silicate_recipe,acrylic_recipe").or("silicate_recipe.eq.true,acrylic_recipe.eq.true");
      const rMap={};
      (cd||[]).forEach(c=>{if(c.silicate_recipe)rMap[c.code+"__силикат"]=true;if(c.acrylic_recipe)rMap[c.code+"__акрил"]=true});
      setRecSync(rMap);
      // Загружаем ВСЕ цвета из базы с флагами рецептур
      const{data:allColors}=await SB.from("colors").select("palette,code,silicate_recipe,acrylic_recipe");
      if(allColors&&allColors.length>0){
        // Обновляем recSync всеми флагами из базы
        setRecSync(prev=>{
          const next={...prev};
          allColors.forEach(c=>{
            if(c.silicate_recipe)next[c.code+"__силикат"]=true;
            if(c.acrylic_recipe)next[c.code+"__акрил"]=true;
          });
          return next;
        });
        setColors(prev=>{
          const existing=new Set(prev.map(c=>c.palette+"|"+c.code));
          const toAdd=allColors.filter(c=>c.palette&&c.code&&!existing.has(c.palette+"|"+c.code));
          if(toAdd.length>0){
            console.log("Adding from DB:",toAdd.length,"colors, palettes:",
              [...new Set(toAdd.map(c=>c.palette))].join(", "));
            return[...prev,...toAdd];
          }
          return prev;
        });
      }
    }catch(e){console.error("Load error:",e)}
    setLoading(false);
  })()},[]);

  const gr=(code,type)=>{
    const k=code+"__"+type;
    if(recSync[k]!==undefined)return recSync[k];
    return type==="силикат"?RS.has(code):RA.has(code);
  };

  const sr=async(code,type,val)=>{
    const col=type==="силикат"?"silicate_recipe":"acrylic_recipe";
    await SB.from("colors").update({[col]:val}).eq("code",code).select().maybeSingle();
    setRecSync(p=>({...p,[code+"__"+type]:val}));
  };

  /* addColor — возвращает промис, обновляет список немедленно */
  const addColor=async(palette,code)=>{
    if(!code.trim()||colors.some(c=>c.palette===palette&&c.code===code.trim()))return false;
    const{error}=await SB.from("colors").insert({palette,code:code.trim(),silicate_recipe:false,acrylic_recipe:false});
    if(error){console.error("addColor error:",error.message);alert("Ошибка сохранения цвета: "+error.message);return false;}
    setColors(p=>[...p,{palette,code:code.trim()}]);
    return true;
  };

  const ao=async(o)=>{
    const row={paint_type:o.paint_type,palette:o.palette,color_code:o.color_code,has_recipe:o.has_recipe,order_type:o.order_type,container_size:o.order_type==="образец"?o.container_size:"20кг",quantity:o.order_type==="образец"?1:parseInt(o.quantity),primer_qty:o.primer_qty||null,object_name:o.object_name||null,facade_area:o.facade_area||null,comment:o.comment||null,desired_date:o.desired_date||null,recipe_version:o.recipe_version||null,status:"new",created_by:o.created_by};
    const{data,error}=await SB.from("orders").insert(row).select().single();
    if(error){alert("Ошибка: "+error.message);return}
    if(data){
      setOrders(p=>[{...data,comments:[],kp:[]},...p]);
      // Авто-создание объекта если указан адрес и объект ещё не существует
      if(o.object_name&&o.object_name.trim()){
        const name=o.object_name.trim();
        const exists=objects.find(obj=>obj.name===name);
        if(!exists){
          try{
            await addObject({name,address:null,contact_name:null,facade_area_total:o.facade_area||null,facade_area_paint:null,paint_type:o.paint_type||null,notes:null,status:"active"});
          }catch(e){console.log("Auto-object:",e.message)}
        }
      }
      setPage("detail");
      setDid(data.id);
    }
  };

  const uo=async(id,u)=>{
    await SB.from("orders").update({...u,updated_at:new Date().toISOString()}).eq("id",id);
    setOrders(p=>p.map(o=>o.id===id?{...o,...u,updated_at:new Date().toISOString()}:o));
    // Автоматически добавляем в рецептуры при отгрузке образца
    if(u.status==="shipped"){
      const order=orders.find(o=>o.id===id);
      if(order&&order.order_type==="образец"){
        try{
          await SB.from("recipes").insert({
            color_code:order.color_code,
            palette:order.palette,
            paint_type:order.paint_type,
            source:"sample",
            confirmed:true
          });
          const col=order.paint_type==="силикат"?"silicate_recipe":"acrylic_recipe";
          await SB.from("colors").update({[col]:true}).eq("code",order.color_code);
          setRecSync(p=>({...p,[order.color_code+"__"+order.paint_type]:true}));
        }catch(e){console.log("Recipe auto-add:",e.message)}
      }
    }
  };

  const ac=async(id,t,a)=>{
    const{data}=await SB.from("order_comments").insert({order_id:id,author:a,text:t}).select().single();
    if(data)setOrders(p=>p.map(o=>o.id===id?{...o,comments:[...(o.comments||[]),data]}:o));
  };

  /* KP CRUD */
  const addKp=async(order_id,kp)=>{
    try{
      const{data}=await SB.from("order_kp").insert({order_id,...kp}).select().single();
      if(data)setOrders(p=>p.map(o=>o.id===order_id?{...o,kp:[...(o.kp||[]),data]}:o));
      return data;
    }catch(e){
      // Таблица ещё не создана — сохраняем локально
      const local={id:"local_"+Date.now(),...kp,order_id,created_at:new Date().toISOString()};
      setOrders(p=>p.map(o=>o.id===order_id?{...o,kp:[...(o.kp||[]),local]}:o));
      return local;
    }
  };
  const delKp=async(order_id,kp_id)=>{
    try{await SB.from("order_kp").delete().eq("id",kp_id)}catch(e){}
    setOrders(p=>p.map(o=>o.id===order_id?{...o,kp:(o.kp||[]).filter(k=>k.id!==kp_id)}:o));
  };

  const delOrder=async(id)=>{
    await SB.from("order_comments").delete().eq("order_id",id);
    await SB.from("orders").delete().eq("id",id);
    setOrders(p=>p.filter(o=>o.id!==id));
    setPage("orders");
  };

  const [recipeFilter,setRecipeFilter]=useState({pal:"Ленинградская",type:"all"});
  const [dbRecipes,setDbRecipes]=useState([]);
  const [objects,setObjects]=useState([]);
  const [samples,setSamples]=useState([]);
  const [oid,setOid]=useState(null);

  useEffect(()=>{(async()=>{
    const{data:od}=await SB.from("objects").select("*").order("created_at",{ascending:false});
    setObjects(od||[]);
    const{data:sd}=await SB.from("samples").select("*").order("created_at",{ascending:false});
    setSamples(sd||[]);
    const{data:rd}=await SB.from("recipes").select("*").order("created_at",{ascending:false});
    const normalizedRd=(rd||[]).map(x=>({...x,paint_type:x.paint_type||x["тип_краски"]||"",palette:x.palette||x["палитра"]||"",composition:x.composition||x["композиция"]||"",source:x.source||x["тип_продукта"]||"sample"}));
    setDbRecipes(normalizedRd);
  })()},[]);

  const addObject=async(o)=>{
    const{data}=await SB.from("objects").insert({...o,created_by:user}).select().single();
    if(data)setObjects(p=>[data,...p]);
    return data;
  };
  const updateObject=async(id,u)=>{
    await SB.from("objects").update({...u,updated_at:new Date().toISOString()}).eq("id",id);
    setObjects(p=>p.map(o=>o.id===id?{...o,...u}:o));
  };

  const addSample=async(s)=>{
    const{data}=await SB.from("samples").insert({...s,created_by:user}).select().single();
    if(data)setSamples(p=>[data,...p]);
    return data;
  };
  const updateSample=async(id,u)=>{
    await SB.from("samples").update({...u,updated_at:new Date().toISOString()}).eq("id",id);
    setSamples(p=>p.map(s=>s.id===id?{...s,...u}:s));
  };
  const deleteSample=async(id)=>{
    await SB.from("samples").delete().eq("id",id);
    setSamples(p=>p.filter(s=>s.id!==id));
  };

  const addDbRecipe=async(r)=>{
    // Нормализуем поля — таблица recipes имеет русские названия колонок
    // Из Supabase: color_code, палитра, тип_краски, тип_продукта, композиция
    const toInsert={
      color_code: r.color_code,
      "палитра": r.palette||r["палитра"]||null,
      "тип_краски": r.paint_type||r["тип_краски"]||null,
      "тип_продукта": r.source||r["тип_продукта"]||"sample",
      "композиция": r.composition||r["композиция"]||null,
    };
    // Проверяем по color_code + тип_краски
    const ptField=r.paint_type||r["тип_краски"];
    const exists=dbRecipes.find(x=>{
      const xPt=x.paint_type||x["тип_краски"];
      return x.color_code===r.color_code&&xPt===ptField;
    });
    if(exists){
      try{
        await SB.from("recipes").update({...toInsert,updated_at:new Date().toISOString()}).eq("id",exists.id);
        setDbRecipes(p=>p.map(x=>x.id===exists.id?{...x,...toInsert,paint_type:ptField,palette:toInsert["палитра"]}:x));
      }catch(e){console.error("Recipe update error:",e)}
      return exists;
    }
    let data=null;
    try{
      const res=await SB.from("recipes").insert(toInsert).select().single();
      data=res.data;
      if(res.error)console.error("Recipe insert error:",res.error);
    }catch(e){console.error("Recipe insert exception:",e)}
    if(data){
      const{data:fresh}=await SB.from("recipes").select("*").order("created_at",{ascending:false});
      // Нормализуем fresh — добавляем JS-поля для удобства
      const normalized=(fresh||[]).map(x=>({...x,paint_type:x.paint_type||x["тип_краски"],palette:x.palette||x["палитра"],composition:x.composition||x["композиция"]}));
      setDbRecipes(normalized);
      const col=ptField==="силикат"?"silicate_recipe":"acrylic_recipe";
      await SB.from("colors").update({[col]:true}).eq("code",r.color_code);
      setRecSync(p=>({...p,[r.color_code+"__"+ptField]:true}));
    }
    return data;
  };

  const gd=id=>{setDid(id);setPage("detail")};
  const [orderTypeFilter,setOrderTypeFilter]=useState("all");
  const nv=pg=>{
    if(pg==="detail_open"){
      // did уже установлен вызывающим кодом
      setPage("detail");
    } else if(pg.startsWith("recipes_pal_")){
      const palName=pg.replace("recipes_pal_","");
      setRecipeFilter({pal:palName,type:"all"});
      setPage("recipes");
    } else if(pg.startsWith("recipes_")){
      const parts=pg.split("_");
      setRecipeFilter({pal:"Ленинградская",type:parts[2]||"all"});
      setPage("recipes");
    } else if(pg.startsWith("obj_")){
      setOid(pg.replace("obj_",""));
      setPage("objdetail");
    } else if(pg.startsWith("orders_status_")){
      // фильтр по статусу из дашборда — просто открываем список
      setOrderTypeFilter("all");
      setPage("orders");
    } else if(pg.startsWith("orders_")){
      setOrderTypeFilter(pg.replace("orders_",""));
      setPage("orders");
    } else if(pg==="neworder_партия"){
      setPage("neworder_партия");
    } else {
      setOrderTypeFilter("all");
      setPage(pg);
    }
  };

  if(loading)return <div style={{fontFamily:FF,background:BG,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{textAlign:"center"}}><div style={{fontSize:22,fontWeight:700,color:RED,marginBottom:6}}>СТМ ТДА</div><div style={{color:G400,fontSize:14}}>Загрузка...</div></div></div>;

  return(
    <div style={{fontFamily:FF,background:BG,minHeight:"100vh",maxWidth:480,margin:"0 auto"}}>
      <div style={{background:W,padding:"14px 20px",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100,boxShadow:"0 1px 0 #e8e8ec"}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <div style={{width:36,height:36,borderRadius:10,background:RED,display:"flex",alignItems:"center",justifyContent:"center"}}>
            <svg width="20" height="20" viewBox="0 0 40 40" fill="none"><path d="M20 6L8 34h6l2.8-7h6.4L25.2 34H32L20 6zm-2.2 18L20 12l2.2 12h-4.4z" fill="#fff"/></svg>
          </div>
          <div>
            <div style={{fontSize:16,fontWeight:700,color:INK,letterSpacing:.5}}>СТМ ТДА</div>
            <div style={{fontSize:9,color:G400,letterSpacing:2,textTransform:"uppercase",fontWeight:500}}>заказы · объекты · рецептуры</div>
          </div>
        </div>
        <div style={{padding:"7px 16px",background:G200,borderRadius:50,fontSize:13,color:G500,fontWeight:600}}>{user}</div>
      </div>
      <div style={{padding:"16px 14px 100px"}}>
        {page==="dash"&&<Dash o={orders} gd={gd} nv={nv} cl={colors} gr={gr} recSync={recSync}/>}
        {(page==="neworder"||page==="neworder_партия")&&<NewOrd cl={colors} gr={gr} ao={ao} user={user} initOrderType={page==="neworder_партия"?"партия":""}/>}
        {page==="orders"&&<OrdList o={orders} gd={gd} initTypeFilter={orderTypeFilter}/>}
        {page==="detail"&&did&&<OrdDet o={orders.find(x=>x.id===did)} uo={uo} ac={ac} del={delOrder} user={user} gr={gr} gb={()=>setPage("orders")} addKp={addKp} delKp={delKp}/>}
        {page==="recipes"&&<RecDB cl={colors} gr={gr} sr={sr} addColor={addColor} initFilter={recipeFilter} dbRecipes={dbRecipes} setDbRecipes={setDbRecipes} orders={orders} objects={objects} samples={samples}/>}
        {page==="objects"&&<ObjList objects={objects} samples={samples} addObject={addObject} updateObject={updateObject} oid={oid} setOid={setOid} nv={nv} user={user}/>}
        {page==="objdetail"&&oid&&<ObjDetail obj={objects.find(x=>x.id===oid)} samples={samples.filter(s=>s.object_id===oid)} addSample={addSample} updateSample={updateSample} deleteSample={deleteSample} addDbRecipe={addDbRecipe} dbRecipes={dbRecipes} user={user} cl={colors} gr={gr} gb={()=>setPage("objects")} ao={ao} orders={orders.filter(o=>o.object_id===oid)}/>}
        {page==="export"&&<Exp o={orders} cl={colors} gr={gr}/>}
      </div>
      <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:480,background:W,boxShadow:"0 -1px 0 #e8e8ec",display:"flex",zIndex:100,paddingBottom:"env(safe-area-inset-bottom)"}}>
        {[["dash","Главная","M3 12l9-9 9 9M9 21V12h6v9"],["neworder","Заказ","M12 5v14M5 12h14"],["orders","Заказы","M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"],["objects","Объекты","M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"],["recipes","Рецептуры","M9 3h6M10 3v7l-6 8a1.5 1.5 0 001.2 2.4h13.6A1.5 1.5 0 0020 18l-6-8V3"],["export","Экспорт","M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"]].map(([id,lb,d])=>{
          const isActive=page===id||(id==="neworder"&&(page==="neworder"||page==="neworder_партия"));
          return <button key={id} onClick={()=>nv(id)} style={{flex:1,padding:"10px 0 8px",border:"none",background:"transparent",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:3,fontFamily:FF}}>
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke={isActive?RED:G300} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d={d}/></svg>
            <span style={{fontSize:10,fontWeight:isActive?700:400,color:isActive?RED:G400}}>{lb}</span>
          </button>;
        })}
      </div>
    </div>
  );
}

/* ═══ DASHBOARD ═══ */
function Dash({o,gd,nv,cl,gr,recSync}){
  const act=o.filter(x=>!["cancelled"].includes(x.status));
  const byStatus=(s)=>o.filter(x=>Array.isArray(s)?s.includes(x.status):x.status===s).length;
  const inWork=o.filter(x=>["inwork","waiting","accepted"].includes(x.status));
  const actSamples=inWork.filter(x=>x.order_type==="образец");
  const actBatch=inWork.filter(x=>x.order_type==="партия");
  const lc=cl.filter(x=>x.palette==="Ленинградская");
  const cc=cl.filter(x=>x.palette==="Caparol");
  const lt=lc.length;const ct=cc.length;
  const ls=Object.keys(recSync).filter(k=>k.endsWith("__силикат")&&recSync[k]&&lc.some(c=>k===c.code+"__силикат")).length;
  const la=Object.keys(recSync).filter(k=>k.endsWith("__акрил")&&recSync[k]&&lc.some(c=>k===c.code+"__акрил")).length;
  const ccCodes=new Set(cc.map(c=>c.code));
  const cs=Object.keys(recSync).filter(k=>k.endsWith("__силикат")&&recSync[k]&&ccCodes.has(k.replace("__силикат",""))).length;
  const ca=Object.keys(recSync).filter(k=>k.endsWith("__акрил")&&recSync[k]&&ccCodes.has(k.replace("__акрил",""))).length;
  const todayD=todayStr();
  const todaySamples=o.filter(x=>x.order_type==="образец"&&!["cancelled"].includes(x.status)&&x.created_at?.slice(0,10)===todayD);
  const todayBatch=o.filter(x=>x.order_type==="партия"&&!["cancelled"].includes(x.status)&&x.created_at?.slice(0,10)===todayD);

  return <div>
    <div style={{display:"flex",gap:8,marginBottom:12}}>
      <button onClick={()=>nv("neworder")} style={{flex:1,padding:"14px 8px",background:RED,color:W,border:"none",borderRadius:14,fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:FF}}>+ Образец</button>
      <button onClick={()=>nv("neworder_партия")} style={{flex:1,padding:"14px 8px",background:NAVY,color:W,border:"none",borderRadius:14,fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:FF}}>+ Партия</button>
      <button onClick={()=>nv("orders")} style={{flex:1,padding:"14px 8px",background:G200,color:G500,border:"none",borderRadius:14,fontSize:14,fontWeight:600,cursor:"pointer",fontFamily:FF}}>Все →</button>
    </div>

    <div style={{...C,marginBottom:8,padding:"10px 12px"}}>
      <div style={{fontSize:10,fontWeight:800,color:G400,textTransform:"uppercase",letterSpacing:1,marginBottom:8}}>Заказы по статусам</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
        {[["Новые",byStatus(["new","accepted"]),RED,"new"],["В работе",byStatus(["inwork","waiting"]),ORG,"inwork"],["Готовы",byStatus("ready"),GRN,"ready"],["Получены",byStatus("received"),"#2563eb","received"],["Отгружены",byStatus("shipped"),G500,"shipped"],["Отменены",byStatus("cancelled"),"#18181b","cancelled"]].map(([l,v,c,k])=>
          <div key={k} onClick={()=>nv("orders_status_"+k)} style={{textAlign:"center",cursor:"pointer",padding:"8px 4px",borderRadius:10,background:v>0?c+"0c":BG,border:`1px solid ${v>0?c+"22":G200}`}}>
            <div style={{fontSize:20,fontWeight:700,color:v>0?c:G400,fontFamily:FF,lineHeight:1}}>{v}</div>
            <div style={{fontSize:9,color:G500,marginTop:3,fontWeight:700,lineHeight:1.2}}>{l}</div>
          </div>
        )}
      </div>
    </div>

    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
      <div onClick={()=>nv("orders_образец")} style={{...C,padding:"12px 14px",cursor:"pointer",borderLeft:`3px solid ${RED}`}}>
        <div style={{fontSize:10,color:RED,fontWeight:800,marginBottom:4}}>🧪 Образцы в работе</div>
        <div style={{fontSize:28,fontWeight:700,color:INK,fontFamily:FF,lineHeight:1}}>{actSamples.length}</div>
        {todaySamples.length>0&&<div style={{fontSize:10,color:GRN,marginTop:4,fontWeight:700}}>+{todaySamples.length} сегодня</div>}
      </div>
      <div onClick={()=>nv("orders_партия")} style={{...C,padding:"12px 14px",cursor:"pointer",borderLeft:`3px solid ${BLU}`}}>
        <div style={{fontSize:10,color:BLU,fontWeight:800,marginBottom:4}}>📦 Партии в работе</div>
        <div style={{fontSize:28,fontWeight:700,color:INK,fontFamily:FF,lineHeight:1}}>{actBatch.length}</div>
        {todayBatch.length>0&&<div style={{fontSize:10,color:GRN,marginTop:4,fontWeight:700}}>+{todayBatch.length} сегодня</div>}
      </div>
    </div>

    <div style={{...C,marginBottom:8,padding:"12px 14px"}}>
      <div style={{fontSize:10,fontWeight:800,color:G400,textTransform:"uppercase",letterSpacing:1,marginBottom:10}}>Рецептуры</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
        <div onClick={()=>nv("recipes_L_силикат")} style={{cursor:"pointer",padding:"10px 12px",borderRadius:12,border:`1.5px solid ${RED}22`,background:RED+"06"}}>
          <div style={{fontSize:10,color:G500,fontWeight:700,marginBottom:4}}>Ленингр. · Силикат</div>
          <div style={{display:"flex",alignItems:"baseline",gap:4}}><span style={{fontSize:22,fontWeight:700,color:RED,fontFamily:FF}}>{ls}</span><span style={{fontSize:11,color:G400}}>/{lt}</span></div>
          <div style={{height:3,borderRadius:2,background:G200,marginTop:6}}><div style={{height:3,borderRadius:2,background:RED,width:`${Math.max(2,lt>0?(ls/lt)*100:0)}%`}}/></div>
        </div>
        <div onClick={()=>nv("recipes_L_акрил")} style={{cursor:"pointer",padding:"10px 12px",borderRadius:12,border:`1.5px solid ${ORG}22`,background:ORG+"06"}}>
          <div style={{fontSize:10,color:G500,fontWeight:700,marginBottom:4}}>Ленингр. · Акрил</div>
          <div style={{display:"flex",alignItems:"baseline",gap:4}}><span style={{fontSize:22,fontWeight:700,color:ORG,fontFamily:FF}}>{la}</span><span style={{fontSize:11,color:G400}}>/{lt}</span></div>
          <div style={{height:3,borderRadius:2,background:G200,marginTop:6}}><div style={{height:3,borderRadius:2,background:ORG,width:`${Math.max(2,lt>0?(la/lt)*100:0)}%`}}/></div>
        </div>
        <div style={{padding:"10px 12px",borderRadius:12,border:`1.5px solid ${G200}`,background:BG}}>
          <div style={{fontSize:10,color:G500,fontWeight:700,marginBottom:4}}>Caparol · Силикат</div>
          <div style={{display:"flex",alignItems:"baseline",gap:4}}><span style={{fontSize:22,fontWeight:700,color:INK,fontFamily:FF}}>{cs}</span><span style={{fontSize:11,color:G400}}>/{ct}</span></div>
          <div style={{height:3,borderRadius:2,background:G200,marginTop:6}}><div style={{height:3,borderRadius:2,background:INK,width:`${Math.max(2,ct>0?(cs/ct)*100:0)}%`}}/></div>
        </div>
        <div style={{padding:"10px 12px",borderRadius:12,border:`1.5px solid ${G200}`,background:BG}}>
          <div style={{fontSize:10,color:G500,fontWeight:700,marginBottom:4}}>Caparol · Акрил</div>
          <div style={{display:"flex",alignItems:"baseline",gap:4}}><span style={{fontSize:22,fontWeight:700,color:INK,fontFamily:FF}}>{ca}</span><span style={{fontSize:11,color:G400}}>/{ct}</span></div>
          <div style={{height:3,borderRadius:2,background:G200,marginTop:6}}><div style={{height:3,borderRadius:2,background:INK,width:`${Math.max(2,ct>0?(ca/ct)*100:0)}%`}}/></div>
        </div>
      </div>
    </div>

    {inWork.length>0&&<div style={{...C,padding:0,overflow:"hidden",marginBottom:8}}>
      <div style={{padding:"10px 14px 8px",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:`1px solid ${G200}`}}>
        <div style={{fontSize:10,fontWeight:800,color:G500,letterSpacing:1,textTransform:"uppercase"}}>В работе</div>
        {inWork.length>4&&<button onClick={()=>nv("orders")} style={{fontSize:11,color:RED,fontWeight:700,background:"none",border:"none",cursor:"pointer",padding:0}}>все {inWork.length} →</button>}
      </div>
      {inWork.slice(0,4).map((x,i)=><div key={x.id} onClick={()=>gd(x.id)} style={{padding:"9px 14px",cursor:"pointer",background:W,borderBottom:i<Math.min(inWork.length,4)-1?`1px solid ${G200}50`:"none",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <div style={{display:"flex",gap:6,alignItems:"center",marginBottom:1}}>
            <span style={{fontWeight:700,fontSize:13,color:INK}}>#{x.order_number}</span>
            <TypeBadge type={x.order_type}/>
          </div>
          <div style={{fontSize:11,color:G500}}>{x.color_code}{x.object_name?` · ${x.object_name}`:""}</div>
          {x.primer_qty&&<div style={{fontSize:10,color:ORG,fontWeight:700}}>+ грунт ×{x.primer_qty}</div>}
        </div>
        <Badge status={x.status}/>
      </div>)}
    </div>}
  </div>
}

/* ═══ NEW ORDER ═══ */
function NewOrd({cl,gr,ao,user,initOrderType=""}){
  const [step,ss]=useState(initOrderType?4:1);const [pt,spt]=useState("");const [pal,spal]=useState("");const [cc,scc]=useState("");const [ot,sot]=useState(initOrderType);const [ct,sct]=useState("");const [qty,sq]=useState(1);const [obj,sobj]=useState("");const [cmt,scmt]=useState("");const [dd,sdd]=useState("");const [area,sar]=useState("");const [srch,ssrch]=useState("");const [rv,srv]=useState("");
  const [primer,setPrimer]=useState(0);
  const allPalettes=[...new Set(cl.map(c=>c.palette))];
  const fl=useMemo(()=>{if(!pal)return[];let f=cl.filter(c=>c.palette===pal);if(srch)f=f.filter(c=>c.code.toLowerCase().includes(srch.toLowerCase()));return f.slice(0,50)},[pal,srch,cl]);
  const sel=cl.find(c=>c.code===cc&&c.palette===pal);
  const hr=sel?gr(sel.code,pt):false;
  const aKg=area?parseFloat(area)*0.4:0;const aBk=aKg?Math.ceil(aKg/20):0;
  const ok=pt&&pal&&cc&&ot&&(ot==="образец"?ct:qty>0);
  const go=()=>ao({paint_type:pt,palette:pal,color_code:cc,has_recipe:hr,order_type:ot,container_size:ot==="образец"?ct:"20кг",quantity:ot==="образец"?1:parseInt(qty),primer_qty:primer>0?primer:null,object_name:obj,comment:cmt,desired_date:dd,facade_area:area?parseFloat(area):null,recipe_version:rv.trim()||null,created_by:user});
  const Sel=({active,onClick,children})=><button onClick={onClick} style={{flex:1,padding:"14px 8px",border:"none",borderRadius:50,background:active?INK:G200,fontWeight:active?700:500,fontSize:14,cursor:"pointer",color:active?W:G500,fontFamily:FF,transition:"all .15s"}}>{children}</button>;

  return <div>
    <Tag>Новый заказ</Tag>
    <div style={{fontSize:24,fontWeight:800,color:INK,marginBottom:20,fontFamily:FF,letterSpacing:-0.5}}>Новый заказ</div>

    <div style={{...C,marginBottom:10}}><Tag>Тип краски</Tag>
      <div style={{display:"flex",gap:10}}>
        <button onClick={()=>{spt("силикат");if(step<2)ss(2)}} style={{flex:1,padding:"16px 8px",border:"none",borderRadius:50,background:pt==="силикат"?INK:G200,fontWeight:700,fontSize:15,cursor:"pointer",color:pt==="силикат"?W:G500,fontFamily:FF,transition:"all .15s"}}>Силикат</button>
        <button onClick={()=>{spt("акрил");if(step<2)ss(2)}} style={{flex:1,padding:"16px 8px",border:"none",borderRadius:50,background:pt==="акрил"?INK:G200,fontWeight:700,fontSize:15,cursor:"pointer",color:pt==="акрил"?W:G500,fontFamily:FF,transition:"all .15s"}}>Акрил</button>
      </div>
    </div>

    {step>=2&&<div style={{...C,marginBottom:10}}><Tag>Палитра</Tag>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
        {allPalettes.map(p=><Sel key={p} active={pal===p} onClick={()=>{spal(p);scc("");ssrch("");if(step<3)ss(3)}}>{p}</Sel>)}
      </div>
    </div>}

    {step>=3&&pal&&<div style={{...C,marginBottom:10}}>
      <Tag>Цвет ({cl.filter(c=>c.palette===pal).length})</Tag>
      <input value={srch} onChange={e=>{ssrch(e.target.value);scc("")}} placeholder="Поиск по коду..." style={{...IB,marginBottom:10}}/>
      {cc&&sel&&<div style={{padding:16,background:hr?GRNL:REDL,borderRadius:14,marginBottom:10,borderLeft:`4px solid ${hr?GRN:RED}`}}>
        <div style={{fontWeight:900,fontSize:17,color:INK,marginBottom:8}}>✓ {cc}</div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
          <RB has={gr(cc,"силикат")} label="Сил"/>
          <RB has={gr(cc,"акрил")} label="Акр"/>
        </div>
        {pt&&!hr&&<div style={{marginTop:8,fontSize:13,color:RED,fontWeight:700}}>⚠ Нет рецептуры — нужен подбор</div>}
        {pt&&hr&&<div style={{marginTop:8,fontSize:13,color:GRN,fontWeight:700}}>✓ Рецептура есть</div>}
      </div>}
      <div style={{maxHeight:200,overflowY:"auto",borderRadius:12,border:`1.5px solid ${G200}`}}>
        {fl.map(c=><div key={c.code} onClick={()=>{scc(c.code);ssrch(c.code);if(step<4)ss(4)}} style={{padding:"11px 16px",cursor:"pointer",background:cc===c.code?REDL:W,borderBottom:`1px solid ${G200}50`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontSize:15,fontWeight:cc===c.code?800:500,color:INK}}>{c.code}</span>
          <div style={{display:"flex",gap:5}}>
            <span style={{width:9,height:9,borderRadius:5,background:gr(c.code,"силикат")?GRN:RED}}/>
            <span style={{width:9,height:9,borderRadius:5,background:gr(c.code,"акрил")?GRN:RED}}/>
          </div>
        </div>)}
        {fl.length===0&&srch&&<div style={{padding:20,color:G400,fontSize:14,textAlign:"center"}}>Не найдено</div>}
      </div>
    </div>}

    {step>=4&&cc&&<div style={{...C,marginBottom:10}}><Tag>Тип заказа</Tag>
      <div style={{display:"flex",gap:8}}>
        <Sel active={ot==="образец"} onClick={()=>{sot("образец");sct("");ss(5)}} c={RED}>🧪 Образец</Sel>
        <Sel active={ot==="партия"} onClick={()=>{sot("партия");sct("20кг");ss(5)}} c={BLU}>📦 Партия</Sel>
      </div>
    </div>}

    {step>=5&&ot&&<div style={{...C,marginBottom:10}}>
      {ot==="образец"?<><Tag>Тара</Tag><div style={{display:"flex",gap:8}}>
        <Sel active={ct==="1.5кг"} onClick={()=>sct("1.5кг")}>1,5 кг (1 л)</Sel>
        <Sel active={ct==="3кг"} onClick={()=>sct("3кг")}>3 кг (2 л)</Sel>
      </div></>:<><Tag>Количество (20 кг)</Tag>
        <div style={{display:"flex",alignItems:"center",gap:20}}>
          <button onClick={()=>sq(Math.max(1,qty-1))} style={{width:52,height:52,borderRadius:14,border:`1.5px solid ${G200}`,background:W,fontSize:24,cursor:"pointer",color:INK,fontFamily:FF}}>−</button>
          <span style={{fontSize:44,fontWeight:700,color:INK,letterSpacing:-1.5,fontFamily:FF}}>{qty}</span>
          <button onClick={()=>sq(qty+1)} style={{width:52,height:52,borderRadius:14,border:`1.5px solid ${G200}`,background:W,fontSize:24,cursor:"pointer",color:INK,fontFamily:FF}}>+</button>
          <span style={{fontSize:15,color:G500,fontWeight:700}}>{qty*20} кг</span>
        </div>
      </>}
    </div>}

    {step>=5&&ot&&<div style={{...C,marginBottom:10}}>
      <Tag>Грунтовка (необязательно)</Tag>
      <div style={{fontSize:12,color:G500,marginBottom:10}}>{pt==="силикат"?"Силиказит / Тиги Сил":"Акриловая грунтовка"} · канистра 10 кг</div>
      <div style={{display:"flex",alignItems:"center",gap:16}}>
        <button onClick={()=>setPrimer(Math.max(0,primer-1))} style={{width:44,height:44,borderRadius:10,border:`1.5px solid ${G200}`,background:W,fontSize:22,cursor:"pointer",color:INK,fontFamily:FF}}>−</button>
        <span style={{fontSize:28,fontWeight:700,color:primer>0?INK:G400,fontFamily:FF,minWidth:32,textAlign:"center"}}>{primer}</span>
        <button onClick={()=>setPrimer(primer+1)} style={{width:44,height:44,borderRadius:10,border:`1.5px solid ${G200}`,background:W,fontSize:22,cursor:"pointer",color:INK,fontFamily:FF}}>+</button>
        {primer>0&&<span style={{fontSize:13,color:G500,fontWeight:700}}>{primer} × 10 кг = {primer*10} кг</span>}
        {primer===0&&<span style={{fontSize:12,color:G400}}>0 канистр</span>}
      </div>
    </div>}

    {step>=5&&<div style={{...C,marginBottom:10}}>
      <Tag>Детали</Tag>
      <input value={obj} onChange={e=>sobj(e.target.value)} placeholder="Объект / адрес" style={{...I,marginBottom:4}}/>
      <input value={dd} onChange={e=>sdd(e.target.value)} type="date" style={{...I,marginBottom:4}}/>
      <div style={{marginBottom:4}}>
        <div style={{fontSize:11,color:G400,fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:6,marginTop:8}}>Версия рецептуры</div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:6}}>
          {["1","1.1","1.2","2","2.1","3"].map(v=><button key={v} onClick={()=>srv(rv===v?"":v)} style={{padding:"7px 14px",borderRadius:50,border:"none",background:rv===v?INK:G200,color:rv===v?W:G500,fontSize:13,fontWeight:rv===v?700:500,cursor:"pointer",fontFamily:FF}}>{v}</button>)}
        </div>
        <input value={rv} onChange={e=>srv(e.target.value)} placeholder="или введите вручную: 1.3, R2..." style={{...I}}/>
      </div>
      <textarea value={cmt} onChange={e=>scmt(e.target.value)} placeholder="Комментарий..." rows={2} style={{...I,resize:"vertical",marginBottom:14}}/>
      <div style={{background:REDL,borderRadius:14,padding:18,borderLeft:`4px solid ${RED}`}}>
        <Tag>Калькулятор</Tag>
        <input value={area} onChange={e=>sar(e.target.value.replace(/[^0-9.]/g,""))} placeholder="Площадь фасада, м²" style={I}/>
        {area&&parseFloat(area)>0&&<div style={{fontSize:15,color:INK,marginTop:12}}>
          <div>Расход (2 слоя): <b style={{color:RED,fontSize:18}}>{aKg.toFixed(0)} кг</b></div>
          <div style={{marginTop:2}}>≈ <b style={{color:RED,fontSize:18}}>{aBk} вёдер</b> × 20 кг</div>
          <div style={{fontSize:11,color:G400,marginTop:6}}>* ориентир, 400 г/м²</div>
        </div>}
      </div>
    </div>}

    {ok&&<button onClick={go} style={{...BTN,background:RED,boxShadow:"none"}}>Создать заказ</button>}
  </div>
}

/* ═══ ORDER CARD ═══ */
function OrdCard({ord,onClick}){
  const isSample=ord.order_type==="образец";
  const typeColor=isSample?RED:NAVY;
  return <div onClick={onClick} style={{...C,cursor:"pointer",padding:"14px",borderRadius:16,display:"flex",flexDirection:"column",gap:8,minHeight:130}}>
    {/* Верхняя строка: номер + статус */}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <span style={{fontWeight:800,fontSize:15,color:INK,letterSpacing:-0.3}}>#{ord.order_number}</span>
      <span style={{fontSize:10,background:SC[ord.status]+"18",color:SC[ord.status],padding:"3px 9px",borderRadius:50,fontWeight:700}}>{ST[ord.status]}</span>
    </div>
    {/* Цвет — главная информация */}
    <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
      <span style={{fontSize:16,fontWeight:800,color:INK,lineHeight:1.2,wordBreak:"break-word"}}>{ord.color_code}</span>
      {ord.recipe_version&&<span style={{fontSize:11,fontWeight:700,color:W,background:ORG,padding:"2px 8px",borderRadius:50}}>v{ord.recipe_version}</span>}
    </div>
    {/* Тип + объём */}
    <div style={{display:"flex",alignItems:"center",gap:6}}>
      <span style={{background:isSample?REDL:BLUL,color:typeColor,padding:"3px 9px",borderRadius:50,fontSize:11,fontWeight:700}}>{isSample?"Образец":"Партия"}</span>
      <span style={{fontSize:12,color:G500,fontWeight:600}}>{isSample?(ord.container_size||""):`${ord.quantity||1}×20кг`}</span>
      {ord.primer_qty&&<span style={{fontSize:11,color:ORG,fontWeight:700}}>+г×{ord.primer_qty}</span>}
    </div>
    {/* Адрес и дата */}
    <div style={{marginTop:"auto"}}>
      {ord.object_name&&<div style={{fontSize:11,color:G400,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",marginBottom:2}}>📍 {ord.object_name}</div>}
      <div style={{fontSize:10,color:G300,fontWeight:500}}>{fmtDateShort(ord.created_at)} · {ord.paint_type}</div>
    </div>
  </div>;
}

/* ═══ ORDER LIST ═══ */
function OrdList({o,gd,initTypeFilter="all"}){
  const [f,sf]=useState("all");
  const [tf,stf]=useState(initTypeFilter);
  const [grouped,setGrouped]=useState(true);
  const statusFiltered=useMemo(()=>{
    if(f==="all")return o;
    if(f==="new")return o.filter(x=>["new","accepted","inwork","waiting","received"].includes(x.status));
    if(f==="ready")return o.filter(x=>x.status==="ready");
    if(f==="shipped")return o.filter(x=>x.status==="shipped");
    return o.filter(x=>x.status===f);
  },[o,f]);
  const list=useMemo(()=>tf==="all"?statusFiltered:statusFiltered.filter(x=>x.order_type===tf),[statusFiltered,tf]);
  // Группировка по дате
  const byDate=useMemo(()=>{
    const m={};
    list.forEach(x=>{const d=x.created_at?x.created_at.slice(0,10):"";if(!m[d])m[d]=[];m[d].push(x)});
    return Object.entries(m).sort((a,b)=>b[0].localeCompare(a[0]));
  },[list]);

  const countLabel=`${list.length} из ${o.length}`;

  return <div>
    <Tag>Заказы</Tag>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:10}}>
      <div style={{fontSize:24,fontWeight:800,color:INK,fontFamily:FF,letterSpacing:-0.5}}>Заказы <span style={{fontSize:14,color:G400,fontWeight:600}}>({countLabel})</span></div>
      <button onClick={()=>setGrouped(!grouped)} style={{fontSize:11,color:grouped?BLU:G500,background:grouped?BLUL:BG,border:"none",borderRadius:8,padding:"4px 10px",cursor:"pointer",fontWeight:700,fontFamily:FF}}>{grouped?"По датам":"Списком"}</button>
    </div>
    {/* Статус */}
    <div style={{display:"flex",gap:6,overflowX:"auto",marginBottom:8}}>
      {[["all","Все"],["new","Новые"],["ready","Готовы"],["shipped","Отгружены"]].map(([k,l])=>{
        const cnt=k==="all"?o.length:k==="new"?o.filter(x=>["new","accepted","inwork","waiting","received"].includes(x.status)).length:o.filter(x=>x.status===k).length;
        return <Pill key={k} active={f===k} onClick={()=>sf(k)}>{l}{cnt>0&&<span style={{fontSize:10,opacity:.7}}> ({cnt})</span>}</Pill>;
      })}
    </div>
    {/* Тип */}
    <div style={{display:"flex",gap:6,marginBottom:12}}>
      {[["all","Все типы"],["образец","🧪 Образцы"],["партия","📦 Партии"]].map(([k,l])=><Pill key={k} active={tf===k} onClick={()=>stf(k)} c={INK}>{l}</Pill>)}
    </div>
    {list.length===0&&<div style={{...C,textAlign:"center",padding:48,color:G400,fontSize:14}}>Нет заказов</div>}

    {grouped
      ?byDate.map(([date,items])=><div key={date} style={{marginBottom:12}}>
          <div style={{fontSize:10,fontWeight:700,color:G400,letterSpacing:.8,textTransform:"uppercase",marginBottom:6,padding:"0 2px"}}>
            {date===todayStr()?"Сегодня":fmtDate(date)} · {items.length}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          {items.map(ord=><OrdCard key={ord.id} ord={ord} onClick={()=>gd(ord.id)}/>)}
          </div>
        </div>)
      :<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        {list.map(ord=><OrdCard key={ord.id} ord={ord} onClick={()=>gd(ord.id)}/>)}
      </div>}
  </div>
}

/* ═══ ORDER DETAIL — COMPACT ═══ */
function OrdDet({o,uo,ac,del,user,gr,gb,addKp,delKp}){
  const [nc,snc]=useState("");
  const today=new Date().toISOString().slice(0,10);
  const [ed,sed]=useState(o?.estimated_date||today);
  const [delConfirm,setDelConfirm]=useState(false);
  const [delReason,setDelReason]=useState("");
  const [showKpForm,setShowKpForm]=useState(false);
  const [kpLink,setKpLink]=useState("");
  const [kpText,setKpText]=useState("");
  const [kpDate,setKpDate]=useState("");
  const [kpSupplier,setKpSupplier]=useState("ВАП");
  const [kpPrice,setKpPrice]=useState("");

  if(!o) return <div style={{...C,textAlign:"center",padding:48,color:G400}}>Заказ не найден</div>;
  const nx=SNEXT[o.status];
  const hr=gr(o.color_code,o.paint_type);
  const isCancelled=o.status==="cancelled";
  const kpList=o.kp||[];

  const handleAddKp=async()=>{
    if(!kpLink.trim()&&!kpText.trim())return;
    await addKp(o.id,{link:kpLink.trim()||null,text:kpText.trim()||null,supplier:kpSupplier.trim()||"ВАП",date:kpDate||null,price_per_kg:kpPrice?parseFloat(kpPrice):null});
    setKpLink("");setKpText("");setKpDate("");setKpPrice("");setShowKpForm(false);
  };

  return <div>
    <button onClick={gb} style={{background:"none",border:"none",color:RED,fontSize:14,fontWeight:800,cursor:"pointer",marginBottom:12,padding:0,fontFamily:FF}}>← К списку</button>

    {/* ═══ ХЕДЕР — компактный, всё важное ═══ */}
    <div style={{...C,marginBottom:8,padding:"14px 16px"}}>
      {/* Строка 1: номер + тип + статус */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:26,fontWeight:700,color:INK,fontFamily:FF}}>#{o.order_number}</span>
          <TypeBadge type={o.order_type}/>
        </div>
        <Badge status={o.status}/>
      </div>
      {/* Строка 2: цвет + палитра + краска */}
      <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:8,flexWrap:"wrap"}}>
        <span style={{fontSize:16,fontWeight:900,color:INK}}>{o.color_code}</span>
        <span style={{fontSize:12,color:G400}}>·</span>
        <span style={{fontSize:13,color:G500}}>{o.palette}</span>
        <span style={{fontSize:12,color:G400}}>·</span>
        <span style={{fontSize:13,fontWeight:700,color:o.paint_type==="силикат"?RED:ORG,textTransform:"uppercase"}}>{o.paint_type}</span>
      </div>
      {/* Строка 3: тара/кол-во + объект + дата */}
      <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:6}}>
        <span style={{fontSize:12,background:BG,borderRadius:6,padding:"3px 8px",color:INK,fontWeight:700}}>
          {o.order_type==="образец"?`📦 обр. ${o.container_size}`:`📦 ${o.quantity}×20кг`}
        </span>
        {o.primer_qty&&<span style={{fontSize:12,background:BG,borderRadius:6,padding:"3px 8px",color:G500}}>грунт {o.primer_qty}×10кг</span>}
        {o.facade_area&&<span style={{fontSize:12,background:BG,borderRadius:6,padding:"3px 8px",color:G500}}>{o.facade_area} м²</span>}
      </div>
      {o.object_name&&<div style={{fontSize:13,color:G500,marginBottom:4}}>📍 {o.object_name}</div>}
      {o.comment&&<div style={{fontSize:12,color:G400,fontStyle:"italic",marginBottom:4}}>"{o.comment}"</div>}
      {/* Рецептура */}
      <div style={{display:"flex",gap:6,marginTop:4,flexWrap:"wrap"}}>
        <RB has={hr} label={o.paint_type==="силикат"?"Силикат":"Акрил"}/>
        <span style={{fontSize:11,color:G400}}>{fmtDateShort(o.created_at)} · {o.created_by}</span>
      </div>
    </div>

    {/* ═══ УПРАВЛЕНИЕ ═══ */}
    <div style={{...C,marginBottom:8,padding:"12px 16px"}}>
      <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:nx||!isCancelled?10:0}}>
        {nx&&<button onClick={()=>uo(o.id,{status:nx})} style={{padding:"10px 20px",background:SC[nx],color:W,border:"none",borderRadius:10,fontSize:14,fontWeight:800,cursor:"pointer",fontFamily:FF,boxShadow:`0 4px 12px ${SC[nx]}44`}}>→ {ST[nx]}</button>}
        {!["cancelled","shipped"].includes(o.status)&&!hr&&o.status!=="waiting"&&
          <button onClick={()=>uo(o.id,{status:"waiting"})} style={{padding:"10px 16px",background:ORGL,color:ORG,border:`1.5px solid ${ORG}44`,borderRadius:10,fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:FF}}>⏳ На подбор</button>}
        {!["cancelled","shipped"].includes(o.status)&&
          <button onClick={()=>uo(o.id,{status:"cancelled"})} style={{padding:"10px 14px",background:REDL,color:RED,border:`1.5px solid ${RED}22`,borderRadius:10,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:FF}}>✕ Отменить</button>}
      </div>
      {/* Дата готовности */}
      <div style={{display:"flex",gap:8,alignItems:"center"}}>
        <input type="date" value={ed} onChange={e=>sed(e.target.value)} style={{...IB,flex:1,padding:"9px 12px",fontSize:13}}/>
        <button onClick={()=>uo(o.id,{estimated_date:ed})} style={{padding:"9px 16px",background:RED,color:W,border:"none",borderRadius:10,fontWeight:800,cursor:"pointer",fontSize:13,fontFamily:FF}}>OK</button>
      </div>
      {o.estimated_date&&<div style={{fontSize:12,color:GRN,marginTop:6,fontWeight:700}}>✓ Готовность: {fmtDate(o.estimated_date)}</div>}
    </div>

    {/* ═══ ДОБАВИТЬ В РЕЦЕПТУРЫ — при получении образца ═══ */}
    {o.order_type==="образец"&&["received","ready"].includes(o.status)&&<RecipeAddBlock order={o} gr={gr} uo={uo}/>}

    {/* ═══ КП — только для партий ═══ */}
    {o.order_type==="партия"&&<div style={{...C,marginBottom:8,padding:0,overflow:"hidden"}}>
      <div style={{padding:"10px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:kpList.length>0||showKpForm?`1px solid ${G200}`:"none"}}>
        <div style={{fontSize:12,fontWeight:800,color:BLU,letterSpacing:.5,textTransform:"uppercase"}}>
          КП {kpList.length>0&&<span style={{background:BLUL,borderRadius:8,padding:"1px 7px",fontSize:11,marginLeft:4}}>{kpList.length}</span>}
        </div>
        <button onClick={()=>setShowKpForm(!showKpForm)} style={{padding:"5px 12px",background:showKpForm?BG:BLU,color:showKpForm?G500:W,border:"none",borderRadius:8,fontSize:12,fontWeight:800,cursor:"pointer",fontFamily:FF}}>
          {showKpForm?"Отмена":"+ КП"}
        </button>
      </div>
      {/* Список КП */}
      {kpList.map((k,i)=><div key={k.id} style={{padding:"10px 16px",borderBottom:i<kpList.length-1?`1px solid ${G200}50`:"none",display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <div style={{flex:1}}>
          {k.link&&<a href={k.link} target="_blank" rel="noopener noreferrer" style={{fontSize:13,color:BLU,fontWeight:700,display:"block",marginBottom:2,wordBreak:"break-all"}}>{k.link.length>40?k.link.slice(0,40)+"...":k.link}</a>}
          {k.text&&<div style={{fontSize:12,color:INK,marginBottom:2}}>{k.text}</div>}
          <div style={{fontSize:11,color:G400}}>
            {k.supplier&&<span>{k.supplier}</span>}
            {k.date&&<span> · {fmtDate(k.date)}</span>}
            {k.price_per_kg&&<span style={{fontWeight:700,color:GRN}}> · {k.price_per_kg} ₽/кг</span>}
          </div>
        </div>
        <button onClick={()=>delKp(o.id,k.id)} style={{background:"none",border:"none",color:G400,cursor:"pointer",padding:"0 0 0 8px",fontSize:16}}>🗑</button>
      </div>)}
      {/* Форма добавления КП */}
      {showKpForm&&<div style={{padding:"12px 16px",background:BLUL}}>
        <input value={kpLink} onChange={e=>setKpLink(e.target.value)} placeholder="Ссылка на КП (Google Docs, PDF...)" style={{...IB,marginBottom:8,padding:"9px 12px",fontSize:13}}/>
        <textarea value={kpText} onChange={e=>setKpText(e.target.value)} placeholder="Текстовое описание или данные КП..." rows={2} style={{...IB,marginBottom:8,resize:"vertical",padding:"9px 12px",fontSize:13}}/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
          <input value={kpSupplier} onChange={e=>setKpSupplier(e.target.value)} placeholder="Поставщик (ВАП)" style={{...IB,padding:"9px 12px",fontSize:13}}/>
          <input value={kpDate} onChange={e=>setKpDate(e.target.value)} type="date" style={{...IB,padding:"9px 12px",fontSize:13}}/>
        </div>
        <div style={{display:"flex",gap:8}}>
          <input value={kpPrice} onChange={e=>setKpPrice(e.target.value.replace(/[^0-9.]/g,""))} placeholder="Цена, ₽/кг" style={{...IB,flex:1,padding:"9px 12px",fontSize:13}}/>
          <button onClick={handleAddKp} style={{padding:"9px 18px",background:BLU,color:W,border:"none",borderRadius:10,fontWeight:800,cursor:"pointer",fontFamily:FF}}>Добавить</button>
        </div>
      </div>}
    </div>}

    {/* ═══ ВОССТАНОВЛЕНИЕ / УДАЛЕНИЕ ═══ */}
    {isCancelled&&<div style={{...C,marginBottom:8,border:`1.5px solid ${RED}22`,padding:"12px 16px"}}>
      <Tag>Отменённый заказ</Tag>
      <button onClick={()=>uo(o.id,{status:"new"})} style={{width:"100%",padding:"10px",background:GRNL,color:GRN,border:`1.5px solid ${GRN}33`,borderRadius:10,fontSize:14,fontWeight:800,cursor:"pointer",fontFamily:FF,marginBottom:8}}>↩ Восстановить</button>
      {!delConfirm?<button onClick={()=>setDelConfirm(true)} style={{width:"100%",padding:"10px",background:REDL,color:RED,border:`1.5px solid ${RED}33`,borderRadius:10,fontSize:14,fontWeight:800,cursor:"pointer",fontFamily:FF}}>🗑 Удалить</button>
      :<div>
        <div style={{fontSize:13,color:INK,fontWeight:700,marginBottom:8}}>Причина:</div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:10}}>
          {["Тест","Дубль","Ошибка","Другое"].map(r=>
            <button key={r} onClick={()=>setDelReason(r)} style={{padding:"6px 12px",borderRadius:8,border:delReason===r?`2px solid ${RED}`:`1.5px solid ${G200}`,background:delReason===r?REDL:W,color:delReason===r?RED:INK,fontSize:12,fontWeight:delReason===r?800:600,cursor:"pointer",fontFamily:FF}}>{r}</button>
          )}
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>{setDelConfirm(false);setDelReason("")}} style={{flex:1,padding:"9px",background:W,border:`1.5px solid ${G200}`,borderRadius:10,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:FF}}>Отмена</button>
          <button onClick={()=>{if(delReason)del(o.id)}} disabled={!delReason} style={{flex:2,padding:"9px",background:delReason?RED:G200,color:W,border:"none",borderRadius:10,fontSize:13,fontWeight:800,cursor:delReason?"pointer":"default",fontFamily:FF}}>Удалить{delReason?` (${delReason})`:""}</button>
        </div>
      </div>}
    </div>}

    {/* ═══ КОММЕНТАРИИ (collapsible) ═══ */}
    <Section title={`Комментарии ${o.comments&&o.comments.length>0?"("+o.comments.length+")":""}`} defaultOpen={false} accent={G500}>
      {o.comments&&o.comments.map(c=><div key={c.id} style={{padding:"10px 12px",background:BG,borderRadius:10,marginBottom:6}}>
        <div style={{fontSize:14,color:INK}}>{c.text}</div>
        <div style={{fontSize:11,color:G400,marginTop:3}}>{c.author} · {fmtDate(c.created_at)}</div>
      </div>)}
      {(!o.comments||!o.comments.length)&&<div style={{fontSize:13,color:G400,marginBottom:10}}>Пока пусто</div>}
      <div style={{display:"flex",gap:8,marginTop:8}}>
        <input value={nc} onChange={e=>snc(e.target.value)} placeholder="Комментарий..." style={{...IB,flex:1,padding:"9px 12px",fontSize:13}} onKeyDown={e=>{if(e.key==="Enter"&&nc.trim()){ac(o.id,nc,user);snc("")}}}/>
        <button onClick={()=>{if(nc.trim()){ac(o.id,nc,user);snc("")}}} style={{padding:"9px 16px",background:RED,color:W,border:"none",borderRadius:10,fontWeight:800,cursor:"pointer",fontFamily:FF}}>↑</button>
      </div>
    </Section>

    {/* ═══ ДОПОЛНИТЕЛЬНЫЕ ДАННЫЕ (collapsible) ═══ */}
    <Section title="Доп. данные" defaultOpen={false} accent={G400}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,fontSize:13}}>
        {[["Палитра",o.palette],["Тара",o.container_size],["Кол-во",o.quantity],["Создал",o.created_by],...(o.recipe_version?[["Версия рецептуры","v"+o.recipe_version]]:[]),...(o.desired_date?[["Желаемый срок",fmtDate(o.desired_date)]]:[])]
          .map(([l,v])=><div key={l}>
            <div style={{fontSize:10,color:G400,fontWeight:800,letterSpacing:.5,textTransform:"uppercase",marginBottom:3}}>{l}</div>
            <div style={{fontWeight:600,color:INK}}>{v||"—"}</div>
          </div>)}
      </div>
    </Section>
  </div>
}

/* ═══ RECIPE ADD BLOCK ═══ */
function RecipeAddBlock({order,gr,uo}){
  const [show,setShow]=useState(false);
  const [silicateDone,setSilicateDone]=useState(()=>gr(order.color_code,"силикат"));
  const [acrylDone,setAcrylDone]=useState(()=>gr(order.color_code,"акрил"));
  const [saving,setSaving]=useState(false);
  const [msg,setMsg]=useState("");

  const handleSave=async()=>{
    setSaving(true);
    // Записываем в таблицу recipes с правильными колонками
    const toInsert={
      color_code:order.color_code,
      "палитра":order.palette,
      "тип_краски":order.paint_type,
      "тип_продукта":"sample",
    };
    try{
      const{data,error}=await SB.from("recipes").insert(toInsert).select().single();
      if(error){
        // Возможно уже есть — попробуем upsert через update
        console.log("Recipe insert error (may already exist):",error.message);
      }
      // Обновляем флаг в colors
      const col=order.paint_type==="силикат"?"silicate_recipe":"acrylic_recipe";
      await SB.from("colors").upsert({palette:order.palette,code:order.color_code,[col]:true},{onConflict:"code"});
      setMsg("✓ Добавлено в рецептуры!");
      if(order.paint_type==="силикат")setSilicateDone(true);
      else setAcrylDone(true);
    }catch(e){
      setMsg("Ошибка: "+e.message);
    }
    setSaving(false);
  };

  const alreadyHas=order.paint_type==="силикат"?silicateDone:acrylDone;

  return <div style={{...C,marginBottom:8,padding:"12px 16px",border:`2px solid ${alreadyHas?GRN+"44":ORG+"44"}`,background:alreadyHas?GRNL:ORGL}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <div>
        <div style={{fontSize:12,fontWeight:800,color:alreadyHas?GRN:ORG}}>
          {alreadyHas?"✓ Рецептура уже в банке":"Добавить в банк рецептур"}
        </div>
        <div style={{fontSize:11,color:G500,marginTop:2}}>{order.color_code} · {order.paint_type} · {order.palette}</div>
      </div>
      {!alreadyHas&&<button onClick={handleSave} disabled={saving} style={{padding:"8px 16px",background:GRN,color:W,border:"none",borderRadius:10,fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:FF}}>
        {saving?"...":"+ В рецептуры"}
      </button>}
    </div>
    {msg&&<div style={{fontSize:12,color:GRN,marginTop:6,fontWeight:700}}>{msg}</div>}
  </div>;
}

/* ═══ EDIT COLOR MODAL ═══ */
function EditColorModal({code,palette,gr,sr,onClose}){
  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.55)",zIndex:200,display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={onClose}>
    <div style={{background:W,borderRadius:"20px 20px 0 0",padding:20,width:"100%",maxWidth:480,boxSizing:"border-box"}} onClick={e=>e.stopPropagation()}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <div>
          <div style={{fontSize:17,fontWeight:800,color:INK}}>{code}</div>
          <div style={{fontSize:12,color:G500,marginTop:2}}>Палитра: {palette}</div>
        </div>
        <button onClick={onClose} style={{background:"none",border:"none",fontSize:22,color:G400,cursor:"pointer",padding:0}}>✕</button>
      </div>
      <div style={{fontSize:12,fontWeight:800,color:G500,textTransform:"uppercase",letterSpacing:.5,marginBottom:10}}>Рецептуры</div>
      <div style={{display:"flex",gap:8,marginBottom:16}}>
        <button onClick={()=>sr(code,"силикат",!gr(code,"силикат"))} style={{flex:1,padding:"12px",borderRadius:12,border:`2px solid ${gr(code,"силикат")?GRN:G200}`,background:gr(code,"силикат")?GRNL:W,color:gr(code,"силикат")?GRN:G500,fontWeight:800,cursor:"pointer",fontFamily:FF,fontSize:13,transition:"all .15s"}}>
          Силикат: {gr(code,"силикат")?"✓ есть":"✕ нет"}
        </button>
        <button onClick={()=>sr(code,"акрил",!gr(code,"акрил"))} style={{flex:1,padding:"12px",borderRadius:12,border:`2px solid ${gr(code,"акрил")?ORG:G200}`,background:gr(code,"акрил")?ORGL:W,color:gr(code,"акрил")?ORG:G500,fontWeight:800,cursor:"pointer",fontFamily:FF,fontSize:13,transition:"all .15s"}}>
          Акрил: {gr(code,"акрил")?"✓ есть":"✕ нет"}
        </button>
      </div>
      <button onClick={onClose} style={{width:"100%",padding:"12px",background:GRN,color:W,border:"none",borderRadius:10,fontSize:14,fontWeight:800,cursor:"pointer",fontFamily:FF}}>Готово</button>
    </div>
  </div>;
}

/* ═══ RECIPES + ADD COLOR ═══ */
function RecDB({cl,gr,sr,addColor,initFilter,dbRecipes,setDbRecipes,orders,objects,samples}){
  const [tab,setTab]=useState("db");
  const [pf,spf]=useState(initFilter?.pal||"Ленинградская");
  const [s,sss]=useState("");
  const [rf,srf]=useState(initFilter?.type==="силикат"?"yes":initFilter?.type==="акрил"?"yes":"all");
  const [pg,spg]=useState(0);const pp=50;
  const [showAdd,setShowAdd]=useState(false);
  const [newCode,setNewCode]=useState("");
  const [addMsg,setAddMsg]=useState("");
  const [dbSearch,setDbSearch]=useState("");
  const [expanded,setExpanded]=useState(null);
  const [colorDetail,setColorDetail]=useState(null);
  const [editColor,setEditColor]=useState(null); // {code, palette} — режим редактирования

  const allPalettes=[...new Set(cl.map(c=>c.palette))];
  const fl=useMemo(()=>{let f=cl;if(pf!=="all")f=f.filter(c=>c.palette===pf);if(s)f=f.filter(c=>c.code.toLowerCase().includes(s.toLowerCase()));if(rf==="yes")f=f.filter(c=>gr(c.code,"силикат")||gr(c.code,"акрил"));if(rf==="no")f=f.filter(c=>!gr(c.code,"силикат")&&!gr(c.code,"акрил"));return f},[cl,pf,s,rf,gr]);
  const pd=fl.slice(pg*pp,(pg+1)*pp);const tp=Math.ceil(fl.length/pp);
  const lc=cl.filter(c=>c.palette==="Ленинградская");

  const [newPalette,setNewPalette]=useState("");
  const [useCustomPalette,setUseCustomPalette]=useState(false);
  const getTargetPalette=()=>useCustomPalette?newPalette.trim():(pf==="all"?"Ленинградская":pf);

  const handleAdd=async()=>{
    const targetPal=getTargetPalette();
    if(!newCode.trim()){setAddMsg("Введите код цвета");return}
    if(useCustomPalette&&!newPalette.trim()){setAddMsg("Введите название палитры");return}
    const ok=await addColor(targetPal,newCode.trim());
    if(ok){setAddMsg("✓ Добавлен: "+newCode.trim()+" в «"+targetPal+"»");setNewCode("");setNewPalette("");setTimeout(()=>setAddMsg(""),3000)}
    else{setAddMsg("Такой цвет уже есть")}
  };

  const filteredDb=useMemo(()=>{
    if(!dbRecipes)return[];
    if(!dbSearch)return dbRecipes;
    return dbRecipes.filter(r=>r.color_code.toLowerCase().includes(dbSearch.toLowerCase())||r.palette?.toLowerCase().includes(dbSearch.toLowerCase()));
  },[dbRecipes,dbSearch]);

  const srcLabel={sample:"Образец",manual:"Вручную",manufacturer:"Производитель"};

  // Счётчики по палитрам для банка рецептур
  const recipesByPalette=useMemo(()=>{
    const m={};
    (dbRecipes||[]).forEach(r=>{
      const p=r.palette||"?";
      if(!m[p])m[p]={total:0,силикат:0,акрил:0};
      m[p].total++;
      if(r.paint_type==="силикат")m[p].силикат++;
      if(r.paint_type==="акрил")m[p].акрил++;
    });
    return m;
  },[dbRecipes]);

  return <div>
    <Tag>Рецептуры</Tag>
    <div style={{fontSize:24,fontWeight:800,color:INK,marginBottom:12,fontFamily:FF,letterSpacing:-0.5}}>Рецептуры</div>

    {/* ═══ COLORS TAB ═══ */}
    <div>
      {/* Статистика по текущей палитре */}
      {pf==="Ленинградская"&&<div style={{fontSize:13,color:G500,marginBottom:10}}>Сил: <b style={{color:RED}}>{lc.filter(c=>gr(c.code,"силикат")).length}</b>/{lc.length} · Акр: <b style={{color:ORG}}>{lc.filter(c=>gr(c.code,"акрил")).length}</b>/{lc.length}</div>}

      {/* Фильтр по палитре — все, включая кастомные */}
      <div style={{display:"flex",gap:6,overflowX:"auto",marginBottom:8,paddingBottom:4}}>
        {allPalettes.map(p=><Pill key={p} active={pf===p} onClick={()=>{spf(p);spg(0)}}>{p}</Pill>)}
      </div>

      <div style={{display:"flex",gap:6,marginBottom:10}}>
        {[["all","Все"],["yes","Есть"],["no","Нет"]].map(([k,l])=><Pill key={k} active={rf===k} onClick={()=>{srf(k);spg(0)}} c={GRN}>{l}</Pill>)}
        <Pill active={showAdd} onClick={()=>setShowAdd(!showAdd)} c={INK}>+ Цвет</Pill>
      </div>

      {/* Форма добавления цвета */}
      {showAdd&&<div style={{...C,marginBottom:10,padding:16,border:`2px dashed ${G200}`}}>
        <div style={{fontSize:12,fontWeight:800,color:INK,marginBottom:10}}>Добавить цвет</div>
        <div style={{display:"flex",gap:6,marginBottom:10}}>
          <button onClick={()=>setUseCustomPalette(false)} style={{flex:1,padding:"7px",border:!useCustomPalette?`2px solid ${RED}`:`1.5px solid ${G200}`,borderRadius:8,background:!useCustomPalette?REDL:W,color:!useCustomPalette?RED:G500,fontSize:12,fontWeight:800,cursor:"pointer",fontFamily:FF}}>В существующую</button>
          <button onClick={()=>setUseCustomPalette(true)} style={{flex:1,padding:"7px",border:useCustomPalette?`2px solid ${RED}`:`1.5px solid ${G200}`,borderRadius:8,background:useCustomPalette?REDL:W,color:useCustomPalette?RED:G500,fontSize:12,fontWeight:800,cursor:"pointer",fontFamily:FF}}>Новая палитра</button>
        </div>
        {useCustomPalette&&<input value={newPalette} onChange={e=>setNewPalette(e.target.value)} placeholder="Название новой палитры (напр. Тиккурила)" style={{...IB,marginBottom:8,padding:"10px 14px"}}/>}
        {!useCustomPalette&&<div style={{fontSize:12,color:G500,marginBottom:8}}>Палитра: <b style={{color:INK}}>{getTargetPalette()}</b></div>}
        <div style={{display:"flex",gap:8}}>
          <input value={newCode} onChange={e=>setNewCode(e.target.value)} placeholder="Код цвета" style={{...IB,flex:1,padding:"10px 14px"}} onKeyDown={e=>{if(e.key==="Enter")handleAdd()}}/>
          <button onClick={handleAdd} style={{padding:"10px 18px",background:INK,color:W,border:"none",borderRadius:10,fontWeight:800,cursor:"pointer",fontFamily:FF}}>+</button>
        </div>
        {addMsg&&<div style={{marginTop:6,fontSize:13,fontWeight:700,color:addMsg.startsWith("✓")?GRN:RED}}>{addMsg}</div>}
      </div>}

      <input value={s} onChange={e=>{sss(e.target.value);spg(0)}} placeholder="Поиск..." style={{...IB,marginBottom:10}}/>
      <div style={{fontSize:11,color:G400,marginBottom:8,fontWeight:700}}>Найдено: {fl.length} · Стр. {pg+1}/{tp||1}</div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:5,marginBottom:10}}>
      {pd.map(c=>{const a=gr(c.code,"силикат"),b=gr(c.code,"акрил");
        return <div key={c.palette+c.code} style={{...C,padding:"9px 10px",cursor:"pointer",borderTop:`2px solid ${a||b?GRN:G200}`}} onClick={()=>setColorDetail({code:c.code,palette:c.palette})}>
          <div style={{fontWeight:600,fontSize:12,color:INK,marginBottom:4,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.code}</div>
          <div style={{display:"flex",gap:3}} onClick={e=>e.stopPropagation()}>
            <button onClick={()=>sr(c.code,"силикат",!a)} style={{flex:1,background:a?GRNL:REDL,color:a?GRN:RED,border:"none",borderRadius:5,padding:"3px 0",fontSize:10,cursor:"pointer",fontWeight:700,fontFamily:FF}}>С{a?"✓":"✕"}</button>
            <button onClick={()=>sr(c.code,"акрил",!b)} style={{flex:1,background:b?GRNL:REDL,color:b?GRN:RED,border:"none",borderRadius:5,padding:"3px 0",fontSize:10,cursor:"pointer",fontWeight:700,fontFamily:FF}}>А{b?"✓":"✕"}</button>
            <button onClick={e=>{e.stopPropagation();setEditColor({code:c.code,palette:c.palette})}} style={{background:BG,color:G400,border:"none",borderRadius:5,padding:"3px 5px",fontSize:10,cursor:"pointer",fontFamily:FF}}>✏</button>
          </div>
        </div>})}
      </div>

      {tp>1&&<div style={{display:"flex",justifyContent:"center",gap:10,marginTop:14}}>
        <button disabled={pg===0} onClick={()=>spg(pg-1)} style={{padding:"10px 22px",borderRadius:10,border:`1.5px solid ${G200}`,background:W,cursor:pg===0?"default":"pointer",opacity:pg===0?.3:1,fontFamily:FF,fontWeight:800}}>←</button>
        <span style={{padding:10,fontSize:14,color:G500,fontWeight:700}}>{pg+1}/{tp}</span>
        <button disabled={pg>=tp-1} onClick={()=>spg(pg+1)} style={{padding:"10px 22px",borderRadius:10,border:`1.5px solid ${G200}`,background:W,cursor:pg>=tp-1?"default":"pointer",opacity:pg>=tp-1?.3:1,fontFamily:FF,fontWeight:800}}>→</button>
      </div>}

      {/* EDIT COLOR MODAL */}
      {editColor&&<EditColorModal code={editColor.code} palette={editColor.palette} gr={gr} sr={sr} onClose={()=>setEditColor(null)}/>}

      {/* COLOR DETAIL MODAL */}
      {colorDetail&&(()=>{
        const code=colorDetail.code;
        const colorOrders=orders.filter(o=>o.color_code===code);
        const colorSamples=samples.filter(s=>s.color_code===code);
        const colorRecipes=dbRecipes?dbRecipes.filter(r=>r.color_code===code):[];
        return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.55)",zIndex:200,display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={()=>setColorDetail(null)}>
          <div style={{background:W,borderRadius:"20px 20px 0 0",padding:20,width:"100%",maxWidth:480,maxHeight:"80vh",overflowY:"auto",boxSizing:"border-box"}} onClick={e=>e.stopPropagation()}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16}}>
              <div>
                <div style={{fontSize:20,fontWeight:800,color:INK,fontFamily:FF}}>{code}</div>
                <div style={{fontSize:12,color:G500,marginTop:2}}>{colorDetail.palette}</div>
              </div>
              <div style={{display:"flex",gap:6}}>
                <button onClick={()=>{setColorDetail(null);setEditColor({code,palette:colorDetail.palette})}} style={{background:BG,border:`1.5px solid ${G200}`,borderRadius:8,padding:"5px 10px",fontSize:12,cursor:"pointer",fontFamily:FF,color:INK}}>✏ Изменить</button>
                <button onClick={()=>setColorDetail(null)} style={{background:"none",border:"none",fontSize:20,color:G400,cursor:"pointer",padding:0}}>✕</button>
              </div>
            </div>
            {/* Рецептуры */}
            {colorRecipes.length>0&&<div style={{marginBottom:16}}>
              <div style={{fontSize:10,fontWeight:800,color:G500,textTransform:"uppercase",letterSpacing:1,marginBottom:8}}>Рецептуры в банке</div>
              {colorRecipes.map(r=><div key={r.id} style={{background:GRNL,borderRadius:10,padding:"10px 12px",marginBottom:6}}>
                <div style={{fontSize:13,fontWeight:800,color:GRN}}>{r.paint_type} · {r.palette}</div>
                {r.composition&&<div style={{fontSize:12,color:INK,marginTop:4}}>{r.composition}</div>}
                <div style={{fontSize:11,color:G400,marginTop:4}}>{fmtDate(r.created_at)}</div>
              </div>)}
            </div>}
            {/* Заказы */}
            {colorOrders.length>0&&<div style={{marginBottom:16}}>
              <div style={{fontSize:10,fontWeight:800,color:G500,textTransform:"uppercase",letterSpacing:1,marginBottom:8}}>Заказы ({colorOrders.length})</div>
              {colorOrders.map(o=><div key={o.id} style={{background:BG,borderRadius:10,padding:"10px 12px",marginBottom:6}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                  <div>
                    <div style={{fontSize:13,fontWeight:800,color:INK}}>#{o.order_number} · {o.paint_type} · {o.order_type==="образец"?`обр. ${o.container_size}`:`${o.quantity}×20кг`}</div>
                    {o.object_name&&<div style={{fontSize:11,color:G500,marginTop:2}}>📍 {o.object_name}</div>}
                  </div>
                  <span style={{background:SC[o.status]+"15",color:SC[o.status],padding:"2px 8px",borderRadius:8,fontSize:10,fontWeight:800}}>{ST[o.status]}</span>
                </div>
                <div style={{fontSize:11,color:G400,marginTop:4}}>{fmtDate(o.created_at)}</div>
              </div>)}
            </div>}
            {colorSamples.length>0&&<div style={{marginBottom:8}}>
              <div style={{fontSize:10,fontWeight:800,color:G500,textTransform:"uppercase",letterSpacing:1,marginBottom:8}}>Образцы ({colorSamples.length})</div>
              {colorSamples.map(s=>{
                const SC3={requested:"#71717a",in_work:"#ea580c",ready:"#d97706",received:"#2563eb",approved:"#16a34a",rejected:"#dc2626"};
                const SL3={requested:"Запрошен",in_work:"В работе",ready:"Готов",received:"Получен",approved:"Согласован",rejected:"Отклонён"};
                return <div key={s.id} style={{background:BG,borderRadius:10,padding:"10px 12px",marginBottom:6}}>
                  <div style={{display:"flex",justifyContent:"space-between"}}>
                    <div style={{fontSize:12,fontWeight:700,color:INK}}>{s.paint_type} · {s.palette}</div>
                    <span style={{background:SC3[s.status]+"15",color:SC3[s.status],padding:"2px 8px",borderRadius:8,fontSize:10,fontWeight:800}}>{SL3[s.status]}</span>
                  </div>
                  <div style={{fontSize:11,color:G400,marginTop:4}}>{fmtDate(s.created_at)}</div>
                </div>})}
            </div>}
            {colorOrders.length===0&&colorSamples.length===0&&colorRecipes.length===0&&
              <div style={{textAlign:"center",padding:"24px 0",color:G400,fontSize:13}}>По этому цвету нет истории</div>}
          </div>
        </div>;
      })()}
    </div>
  </div>
}


/* ═══ OBJECTS LIST ═══ */
function ObjList({objects,samples,addObject,updateObject,oid,setOid,nv,user}){
  const [showForm,setShowForm]=useState(false);
  const [name,setName]=useState("");const [addr,setAddr]=useState("");const [contact,setContact]=useState("");
  const [submitting,setSubmitting]=useState(false);
  const [areaTotal,setAreaTotal]=useState("");
  const [areaPaint,setAreaPaint]=useState("");
  const [paintType,setPaintTypeO]=useState("");
  const [objNotes,setObjNotes]=useState("");

  const handleAdd=async()=>{
    if(!name.trim())return;
    setSubmitting(true);
    await addObject({name:name.trim(),address:addr.trim()||null,contact_name:contact.trim()||null,facade_area_total:areaTotal?parseFloat(areaTotal):null,facade_area_paint:areaPaint?parseFloat(areaPaint):null,paint_type:paintType||null,notes:objNotes.trim()||null,status:"active"});
    setName("");setAddr("");setContact("");setAreaTotal("");setAreaPaint("");setPaintTypeO("");setObjNotes("");
    setShowForm(false);setSubmitting(false);
  };

  const statusColor={active:GRN,completed:G400,paused:ORG};
  const statusLabel={active:"Активный",completed:"Завершён",paused:"На паузе"};

  return <div>
    <Tag>Объекты</Tag>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:14}}>
      <div style={{fontSize:24,fontWeight:800,color:INK,fontFamily:FF,letterSpacing:-0.5}}>Объекты</div>
      <button onClick={()=>setShowForm(!showForm)} style={{padding:"8px 16px",background:showForm?BG:RED,color:showForm?G500:W,border:showForm?`1.5px solid ${G200}`:"none",borderRadius:10,fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:FF}}>
        {showForm?"Отмена":"+ Объект"}
      </button>
    </div>

    {showForm&&<div style={{...C,marginBottom:12,padding:16,border:`2px dashed ${RED}44`}}>
      <Tag>Новый объект</Tag>
      <input value={name} onChange={e=>setName(e.target.value)} placeholder="Адрес объекта *" style={{...IB,marginBottom:8}}/>
      <input value={contact} onChange={e=>setContact(e.target.value)} placeholder="Архитектор / заказчик" style={{...IB,marginBottom:8}}/>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
        <input value={areaTotal} onChange={e=>setAreaTotal(e.target.value.replace(/[^0-9.]/g,""))} placeholder="Площадь фасада, м²" style={{...IB,padding:"10px 12px"}}/>
        <input value={areaPaint} onChange={e=>setAreaPaint(e.target.value.replace(/[^0-9.]/g,""))} placeholder="Под окраску, м²" style={{...IB,padding:"10px 12px"}}/>
      </div>
      <div style={{display:"flex",gap:6,marginBottom:8}}>
        {["силикат","акрил","оба"].map(t=><button key={t} onClick={()=>setPaintTypeO(paintType===t?"":t)} style={{flex:1,padding:"8px 4px",border:paintType===t?`2px solid ${t==="акрил"?ORG:RED}`:`1.5px solid ${G200}`,borderRadius:8,background:paintType===t?t==="акрил"?ORGL:REDL:W,color:paintType===t?t==="акрил"?ORG:RED:G500,fontSize:12,fontWeight:800,cursor:"pointer",fontFamily:FF,textTransform:"uppercase"}}>{t}</button>)}
      </div>
      <textarea value={objNotes} onChange={e=>setObjNotes(e.target.value)} placeholder="Примечания..." rows={2} style={{...IB,marginBottom:10,resize:"vertical"}}/>
      <button onClick={handleAdd} disabled={!name.trim()||submitting} style={{width:"100%",padding:"12px",background:name.trim()?`linear-gradient(135deg,${RED},#b91c1c)`:G200,color:W,border:"none",borderRadius:10,fontSize:14,fontWeight:800,cursor:name.trim()?"pointer":"default",fontFamily:FF}}>
        {submitting?"Создаём...":"Создать объект"}
      </button>
    </div>}

    {objects.length===0&&!showForm&&<div style={{...C,textAlign:"center",padding:48,color:G400}}>
      <div style={{fontSize:32,marginBottom:8}}>🏗</div>
      <div style={{fontSize:14}}>Объектов пока нет</div>
    </div>}

    {objects.map(obj=>{
      const objSamples=samples.filter(s=>s.object_id===obj.id);
      const approved=objSamples.filter(s=>s.status==="approved").length;
      const total=objSamples.length;
      return <div key={obj.id} onClick={()=>nv("obj_"+obj.id)} style={{...C,marginBottom:8,cursor:"pointer",borderLeft:`4px solid ${statusColor[obj.status]||G400}`,padding:16}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4}}>
          <div style={{fontWeight:800,fontSize:15,color:INK}}>{obj.name}</div>
          <span style={{background:statusColor[obj.status]+"15",color:statusColor[obj.status],padding:"3px 10px",borderRadius:12,fontSize:11,fontWeight:800}}>{statusLabel[obj.status]}</span>
        </div>
        {obj.address&&<div style={{fontSize:12,color:G500,marginBottom:2}}>📍 {obj.address}</div>}
        {obj.contact_name&&<div style={{fontSize:12,color:G400,marginBottom:2}}>👤 {obj.contact_name}</div>}
        <div style={{display:"flex",gap:10,flexWrap:"wrap",marginTop:4}}>
          {obj.facade_area_total&&<span style={{fontSize:11,color:G500}}>Фасад: <b style={{color:INK}}>{obj.facade_area_total} м²</b></span>}
          {obj.facade_area_paint&&<span style={{fontSize:11,color:G500}}>Окраска: <b style={{color:INK}}>{obj.facade_area_paint} м²</b></span>}
          {obj.paint_type&&<span style={{fontSize:11,background:obj.paint_type==="акрил"?ORGL:REDL,color:obj.paint_type==="акрил"?ORG:RED,padding:"2px 8px",borderRadius:6,fontWeight:800,textTransform:"uppercase"}}>{obj.paint_type}</span>}
        </div>
        {total>0&&<div style={{marginTop:6,fontSize:12,color:G500}}>Образцы: <b style={{color:INK}}>{total}</b> · Согласовано: <b style={{color:GRN}}>{approved}</b></div>}
      </div>
    })}
  </div>
}

/* ═══ OBJECT DETAIL ═══ */
function ObjDetail({obj,samples,addSample,updateSample,deleteSample,addDbRecipe,dbRecipes,user,cl,gr,gb,ao,orders}){
  const [showSampleForm,setShowSampleForm]=useState(false);
  const [pal,setPal]=useState("Ленинградская");
  const [cc,setCc]=useState("");const [srch,setSrch]=useState("");
  const [pt,setPt]=useState("силикат");
  const [notes,setNotes]=useState("");
  const [submitting,setSubmitting]=useState(false);

  const allPalettes=[...new Set(cl.map(c=>c.palette))];
  const fl=useMemo(()=>{
    let f=cl.filter(c=>c.palette===pal);
    if(srch)f=f.filter(c=>c.code.toLowerCase().includes(srch.toLowerCase()));
    return f.slice(0,40);
  },[cl,pal,srch]);

  const SC2={requested:"#71717a",in_work:"#ea580c",ready:"#d97706",received:"#2563eb",approved:"#16a34a",rejected:"#dc2626"};
  const SL={requested:"Запрошен",in_work:"В работе",ready:"Готов",received:"Получен",approved:"Согласован",rejected:"Отклонён"};
  const SN={requested:"in_work",in_work:"ready",ready:"received",received:"approved"};

  const handleAddSample=async()=>{
    if(!cc)return;
    setSubmitting(true);
    await addSample({object_id:obj.id,color_code:cc,palette:pal,paint_type:pt,status:"requested",notes:notes||null});
    setCc("");setSrch("");setNotes("");setShowSampleForm(false);setSubmitting(false);
  };

  const handleReceive=async(s)=>{
    await updateSample(s.id,{status:"received",received_by:user,received_at:new Date().toISOString()});
  };

  const [approveModal,setApproveModal]=useState(null);
  const [composition,setComposition]=useState("");
  const [approving,setApproving]=useState(false);

  const handleApprove=async()=>{
    const s=approveModal;
    if(!s)return;
    setApproving(true);
    await updateSample(s.id,{status:"approved",approved_at:new Date().toISOString()});
    await addDbRecipe({color_code:s.color_code,palette:s.palette,paint_type:s.paint_type,source:"sample",source_sample_id:s.id,source_object_id:s.object_id,composition:composition.trim()||null,confirmed:true,notes:s.notes||null});
    setComposition("");setApproveModal(null);setApproving(false);
  };

  if(!obj)return <div style={{...C,padding:48,textAlign:"center",color:G400}}>Объект не найден</div>;

  const approved=samples.filter(s=>s.status==="approved");
  const active=samples.filter(s=>!["approved","rejected"].includes(s.status));

  return <div>
    <button onClick={gb} style={{background:"none",border:"none",color:RED,fontSize:14,fontWeight:800,cursor:"pointer",marginBottom:12,padding:0,fontFamily:FF}}>← К объектам</button>

    {/* OBJ INFO */}
    <div style={{...C,marginBottom:10}}>
      <div style={{fontSize:22,fontWeight:700,color:INK,fontFamily:FF,marginBottom:8,lineHeight:1.2}}>{obj.name}</div>
      {obj.address&&<div style={{fontSize:13,color:G500,marginBottom:3}}>📍 {obj.address}</div>}
      {obj.contact_name&&<div style={{fontSize:13,color:G500,marginBottom:3}}>👤 {obj.contact_name}</div>}
      <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:obj.notes?8:0,marginTop:6}}>
        {obj.facade_area_total&&<div style={{background:BG,borderRadius:8,padding:"6px 12px"}}><div style={{fontSize:9,color:G400,fontWeight:800,textTransform:"uppercase",letterSpacing:.5}}>Фасад всего</div><div style={{fontSize:15,fontWeight:700,color:INK}}>{obj.facade_area_total} м²</div></div>}
        {obj.facade_area_paint&&<div style={{background:BG,borderRadius:8,padding:"6px 12px"}}><div style={{fontSize:9,color:G400,fontWeight:800,textTransform:"uppercase",letterSpacing:.5}}>Под окраску</div><div style={{fontSize:15,fontWeight:700,color:INK}}>{obj.facade_area_paint} м²</div></div>}
        {obj.paint_type&&<div style={{background:obj.paint_type==="акрил"?ORGL:REDL,borderRadius:8,padding:"6px 12px"}}><div style={{fontSize:9,color:G400,fontWeight:800,textTransform:"uppercase",letterSpacing:.5}}>Тип</div><div style={{fontSize:13,fontWeight:800,color:obj.paint_type==="акрил"?ORG:RED,textTransform:"uppercase"}}>{obj.paint_type}</div></div>}
      </div>
      {obj.notes&&<div style={{fontSize:13,color:G500,background:BG,borderRadius:8,padding:"8px 12px",marginBottom:4}}>{obj.notes}</div>}
      <div style={{marginTop:10,display:"flex",gap:8}}>
        <div style={{flex:1,background:REDL,borderRadius:10,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:22,fontWeight:700,color:RED,fontFamily:FF}}>{samples.length}</div><div style={{fontSize:10,color:G500,fontWeight:700,marginTop:2}}>Образцов</div></div>
        <div style={{flex:1,background:GRNL,borderRadius:10,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:22,fontWeight:700,color:GRN,fontFamily:FF}}>{approved.length}</div><div style={{fontSize:10,color:G500,fontWeight:700,marginTop:2}}>Согласовано</div></div>
        <div style={{flex:1,background:BG,borderRadius:10,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:22,fontWeight:700,color:INK,fontFamily:FF}}>{orders.length}</div><div style={{fontSize:10,color:G500,fontWeight:700,marginTop:2}}>Заказов</div></div>
      </div>
    </div>

    {/* ADD SAMPLE */}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
      <Tag>Образцы</Tag>
      <button onClick={()=>setShowSampleForm(!showSampleForm)} style={{padding:"6px 14px",background:showSampleForm?BG:INK,color:showSampleForm?G500:W,border:showSampleForm?`1.5px solid ${G200}`:"none",borderRadius:8,fontSize:12,fontWeight:800,cursor:"pointer",fontFamily:FF,marginBottom:8}}>
        {showSampleForm?"Отмена":"+ Образец"}
      </button>
    </div>

    {showSampleForm&&<div style={{...C,marginBottom:10,padding:14,border:`2px dashed ${G200}`}}>
      <div style={{display:"flex",gap:6,marginBottom:10,overflowX:"auto"}}>
        {allPalettes.map(p=><button key={p} onClick={()=>{setPal(p);setCc("");setSrch("")}} style={{flexShrink:0,padding:"8px 10px",border:pal===p?`2px solid ${RED}`:`1.5px solid ${G200}`,borderRadius:8,background:pal===p?REDL:W,color:pal===p?RED:G500,fontSize:11,fontWeight:pal===p?800:600,cursor:"pointer",fontFamily:FF}}>{p==="Ленинградская"?"Ленингр.":p}</button>)}
      </div>
      <div style={{display:"flex",gap:6,marginBottom:10}}>
        {["силикат","акрил"].map(t=><button key={t} onClick={()=>setPt(t)} style={{flex:1,padding:"8px",border:pt===t?`2px solid ${t==="силикат"?RED:ORG}`:`1.5px solid ${G200}`,borderRadius:8,background:pt===t?t==="силикат"?REDL:ORGL:W,color:pt===t?t==="силикат"?RED:ORG:G500,fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:FF,textTransform:"uppercase"}}>{t}</button>)}
      </div>
      <input value={srch} onChange={e=>{setSrch(e.target.value);setCc("")}} placeholder="Поиск цвета..." style={{...IB,marginBottom:6,padding:"10px 14px"}}/>
      {cc&&<div style={{padding:"8px 12px",background:GRNL,borderRadius:8,marginBottom:6,fontSize:13,fontWeight:700,color:GRN}}>✓ {cc}</div>}
      <div style={{maxHeight:160,overflowY:"auto",borderRadius:10,border:`1.5px solid ${G200}`,marginBottom:8}}>
        {fl.map(c=><div key={c.code} onClick={()=>{setCc(c.code);setSrch(c.code)}} style={{padding:"9px 12px",cursor:"pointer",background:cc===c.code?REDL:W,borderBottom:`1px solid ${G200}50`,fontSize:13,fontWeight:cc===c.code?800:500,color:INK}}>{c.code}</div>)}
      </div>
      <input value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Примечание (необязательно)" style={{...IB,marginBottom:10,padding:"10px 14px"}}/>
      <button onClick={handleAddSample} disabled={!cc||submitting} style={{width:"100%",padding:"11px",background:cc?`linear-gradient(135deg,${RED},#b91c1c)`:G200,color:W,border:"none",borderRadius:10,fontSize:14,fontWeight:800,cursor:cc?"pointer":"default",fontFamily:FF}}>
        {submitting?"Добавляем...":"Добавить образец"}
      </button>
    </div>}

    {/* SAMPLE LIST */}
    {samples.length===0&&<div style={{...C,textAlign:"center",padding:32,color:G400,fontSize:13}}>Образцов пока нет</div>}
    {samples.map(s=><div key={s.id} style={{...C,marginBottom:6,padding:14,borderLeft:`4px solid ${SC2[s.status]||G400}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
        <div>
          <div style={{fontWeight:800,fontSize:14,color:INK}}>{s.color_code}</div>
          <div style={{fontSize:11,color:G500,marginTop:1}}>{s.palette} · {s.paint_type}</div>
        </div>
        <span style={{background:SC2[s.status]+"15",color:SC2[s.status],padding:"3px 10px",borderRadius:12,fontSize:11,fontWeight:800}}>{SL[s.status]}</span>
      </div>
      {s.notes&&<div style={{fontSize:12,color:G400,marginBottom:6}}>{s.notes}</div>}
      {s.received_by&&<div style={{fontSize:11,color:G400,marginBottom:4}}>Получил: {s.received_by}</div>}
      <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
        {SN[s.status]&&s.status!=="received"&&<button onClick={()=>updateSample(s.id,{status:SN[s.status]})} style={{padding:"6px 12px",background:SC2[SN[s.status]]+"15",color:SC2[SN[s.status]],border:`1px solid ${SC2[SN[s.status]]}33`,borderRadius:8,fontSize:12,fontWeight:800,cursor:"pointer",fontFamily:FF}}>→ {SL[SN[s.status]]}</button>}
        {s.status==="ready"&&<button onClick={()=>handleReceive(s)} style={{padding:"6px 12px",background:"#eff6ff",color:"#2563eb",border:"1px solid #2563eb33",borderRadius:8,fontSize:12,fontWeight:800,cursor:"pointer",fontFamily:FF}}>Получила я</button>}
        {s.status==="received"&&<button onClick={()=>{setApproveModal(s);setComposition("")}} style={{padding:"6px 12px",background:GRNL,color:GRN,border:`1px solid ${GRN}33`,borderRadius:8,fontSize:12,fontWeight:800,cursor:"pointer",fontFamily:FF}}>✓ Согласовать</button>}
        {s.status==="received"&&<button onClick={()=>updateSample(s.id,{status:"rejected"})} style={{padding:"6px 12px",background:REDL,color:RED,border:`1px solid ${RED}33`,borderRadius:8,fontSize:12,fontWeight:800,cursor:"pointer",fontFamily:FF}}>✕ Отклонить</button>}
        {["requested","rejected"].includes(s.status)&&<button onClick={()=>deleteSample(s.id)} style={{padding:"6px 10px",background:BG,color:G400,border:`1px solid ${G200}`,borderRadius:8,fontSize:11,cursor:"pointer",fontFamily:FF}}>🗑</button>}
      </div>
    </div>)}

    {/* APPROVE MODAL */}
    {approveModal&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.5)",zIndex:200,display:"flex",alignItems:"flex-end",justifyContent:"center"}}>
      <div style={{background:W,borderRadius:"20px 20px 0 0",padding:24,width:"100%",maxWidth:480,boxSizing:"border-box"}}>
        <div style={{fontSize:16,fontWeight:800,color:INK,marginBottom:4}}>Согласовать образец</div>
        <div style={{fontSize:13,color:G500,marginBottom:14}}>{approveModal.color_code} · {approveModal.paint_type} · {approveModal.palette}</div>
        <div style={{fontSize:12,fontWeight:700,color:G500,marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Состав рецептуры (необязательно)</div>
        <textarea value={composition} onChange={e=>setComposition(e.target.value)} placeholder="Например: база А 85%, пигмент Л-5 15%..." rows={3} style={{...IB,marginBottom:14,resize:"vertical"}}/>
        <div style={{fontSize:11,color:G400,marginBottom:14}}>Цвет автоматически попадёт в банк рецептур.</div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>{setApproveModal(null);setComposition("")}} style={{flex:1,padding:"12px",background:W,border:`1.5px solid ${G200}`,borderRadius:10,fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:FF}}>Отмена</button>
          <button onClick={handleApprove} disabled={approving} style={{flex:2,padding:"12px",background:GRN,color:W,border:"none",borderRadius:10,fontSize:14,fontWeight:800,cursor:"pointer",fontFamily:FF,boxShadow:`0 4px 12px ${GRN}44`}}>
            {approving?"Сохраняем...":"✓ Согласовать + в банк"}
          </button>
        </div>
      </div>
    </div>}

    {/* AGREED COLORS */}
    {approved.length>0&&<div style={{...C,marginTop:12,padding:14,border:`2px solid ${GRN}33`,background:GRNL}}>
      <div style={{fontSize:12,fontWeight:800,color:GRN,marginBottom:8}}>✓ СОГЛАСОВАННЫЕ ({approved.length})</div>
      {approved.map(s=><div key={s.id} style={{fontSize:13,color:INK,fontWeight:700,marginBottom:4}}>· {s.color_code} ({s.paint_type})</div>)}
    </div>}
  </div>
}

/* ═══ EXPORT ═══ */
function Exp({o,cl,gr}){
  const [sortField,setSortField]=useState("created_at");
  const [sortDir,setSortDir]=useState("desc");
  const [filterStatus,setFilterStatus]=useState("all");
  const [filterType,setFilterType]=useState("all");

  const sortOrders=(data)=>{
    const sorted=[...data].sort((a,b)=>{
      let va=a[sortField]||"",vb=b[sortField]||"";
      if(sortField==="order_number"){va=Number(va);vb=Number(vb);}
      if(sortField==="quantity"){va=Number(va)||0;vb=Number(vb)||0;}
      if(va<vb)return sortDir==="asc"?-1:1;
      if(va>vb)return sortDir==="asc"?1:-1;
      return 0;
    });
    return sorted;
  };

  const filteredOrders=useMemo(()=>{
    let d=o;
    if(filterStatus!=="all")d=d.filter(x=>x.status===filterStatus);
    if(filterType!=="all")d=d.filter(x=>x.order_type===filterType);
    return sortOrders(d);
  },[o,filterStatus,filterType,sortField,sortDir]);

  const xo=()=>{
    const r=filteredOrders.map(x=>({"№":x.order_number,"Дата":fmtDate(x.created_at),"Статус":ST[x.status],"Цвет":x.color_code,"Палитра":x.palette,"Краска":x.paint_type,"Тип":x.order_type,"Тара":x.container_size,"Кол-во":x.quantity||1,"Грунт (канистр)":x.primer_qty||0,"Объект":x.object_name||"","Площадь м²":x.facade_area||"","Рецептура":x.has_recipe?"Есть":"Нет","Создал":x.created_by,"Желаемый срок":x.desired_date?fmtDate(x.desired_date):""}));
    const ws=XLSX.utils.json_to_sheet(r);
    ws["!cols"]=[{wch:6},{wch:12},{wch:14},{wch:22},{wch:14},{wch:10},{wch:8},{wch:8},{wch:8},{wch:8},{wch:24},{wch:10},{wch:10},{wch:10},{wch:14}];
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Заказы");
    XLSX.writeFile(wb,`stm_orders_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  const xr=pf=>{
    let d=cl;if(pf!=="all")d=d.filter(c=>c.palette===pf);
    const r=d.map(c=>({"Палитра":c.palette,"Код":c.code,"Силикат":gr(c.code,"силикат")?"Есть":"Нет","Акрил":gr(c.code,"акрил")?"Есть":"Нет"}));
    const ws=XLSX.utils.json_to_sheet(r);ws["!cols"]=[{wch:16},{wch:22},{wch:10},{wch:10}];
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Рецептуры");
    XLSX.writeFile(wb,"stm_tda_recipes.xlsx");
  };

  const SortBtn=({field,label})=><button onClick={()=>{if(sortField===field)setSortDir(d=>d==="asc"?"desc":"asc");else{setSortField(field);setSortDir("desc")}}} style={{padding:"5px 10px",border:`1.5px solid ${sortField===field?RED:G200}`,borderRadius:7,background:sortField===field?REDL:W,color:sortField===field?RED:G500,fontSize:11,fontWeight:sortField===field?700:500,cursor:"pointer",fontFamily:FF}}>{label}{sortField===field?(sortDir==="desc"?" ↓":" ↑"):""}</button>;

  const allPalettes=[...new Set(cl.map(c=>c.palette))];

  return <div>
    <div style={{fontSize:22,fontWeight:700,color:INK,marginBottom:14,fontFamily:FF}}>Экспорт</div>

    {/* ЗАКАЗЫ */}
    <div style={{...C,marginBottom:10}}>
      <div style={{fontSize:11,fontWeight:700,color:G400,textTransform:"uppercase",letterSpacing:.8,marginBottom:10}}>Заказы · {filteredOrders.length} шт.</div>

      {/* Фильтры */}
      <div style={{marginBottom:8}}>
        <div style={{fontSize:10,color:G400,marginBottom:4}}>Статус</div>
        <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
          {[["all","Все"],["new","Новые"],["inwork","В работе"],["ready","Готовы"],["shipped","Отгружены"],["cancelled","Отменены"]].map(([k,l])=>
            <button key={k} onClick={()=>setFilterStatus(k)} style={{padding:"4px 9px",border:`1.5px solid ${filterStatus===k?RED:G200}`,borderRadius:6,background:filterStatus===k?REDL:W,color:filterStatus===k?RED:G500,fontSize:11,fontWeight:filterStatus===k?700:500,cursor:"pointer",fontFamily:FF}}>{l}</button>
          )}
        </div>
      </div>
      <div style={{marginBottom:10}}>
        <div style={{fontSize:10,color:G400,marginBottom:4}}>Тип</div>
        <div style={{display:"flex",gap:4}}>
          {[["all","Все"],["образец","Образцы"],["партия","Партии"]].map(([k,l])=>
            <button key={k} onClick={()=>setFilterType(k)} style={{padding:"4px 9px",border:`1.5px solid ${filterType===k?BLU:G200}`,borderRadius:6,background:filterType===k?BLUL:W,color:filterType===k?BLU:G500,fontSize:11,fontWeight:filterType===k?700:500,cursor:"pointer",fontFamily:FF}}>{l}</button>
          )}
        </div>
      </div>

      {/* Сортировка */}
      <div style={{marginBottom:12}}>
        <div style={{fontSize:10,color:G400,marginBottom:4}}>Сортировка</div>
        <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
          <SortBtn field="created_at" label="По дате"/>
          <SortBtn field="order_number" label="По номеру"/>
          <SortBtn field="color_code" label="По цвету"/>
          <SortBtn field="object_name" label="По объекту"/>
          <SortBtn field="quantity" label="По кол-ву"/>
        </div>
      </div>

      <button onClick={xo} style={{width:"100%",padding:"12px",background:GRN,color:W,border:"none",borderRadius:10,fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:FF,boxShadow:`0 2px 8px ${GRN}40`}}>
        ↓ Скачать .xlsx ({filteredOrders.length} строк)
      </button>
    </div>

    {/* РЕЦЕПТУРЫ */}
    <div style={{...C}}>
      <div style={{fontSize:11,fontWeight:700,color:G400,textTransform:"uppercase",letterSpacing:.8,marginBottom:10}}>Рецептуры</div>
      <div style={{display:"flex",flexDirection:"column",gap:6}}>
        {allPalettes.map(p=><div key={p} onClick={()=>xr(p)} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 12px",borderRadius:8,border:`1.5px solid ${G200}`,cursor:"pointer",background:BG}}>
          <div>
            <div style={{fontSize:13,fontWeight:600,color:INK}}>{p}</div>
            <div style={{fontSize:10,color:G400}}>{cl.filter(c=>c.palette===p).length} цветов</div>
          </div>
          <span style={{fontSize:12,color:RED,fontWeight:700}}>↓ xlsx</span>
        </div>)}
        <div onClick={()=>xr("all")} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 12px",borderRadius:8,border:`1.5px solid ${RED}33`,cursor:"pointer",background:REDL}}>
          <div style={{fontSize:13,fontWeight:700,color:RED}}>Все палитры</div>
          <span style={{fontSize:12,color:RED,fontWeight:700}}>↓ xlsx</span>
        </div>
      </div>
    </div>
  </div>
}
