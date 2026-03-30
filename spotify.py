"""
Spotify Artist Counter
Counts how many songs you have from each artist in your liked songs
Handles large libraries (16500+) with rate limiting and error recovery
"""

import spotipy
from spotipy.oauth2 import SpotifyOAuth
import time
import json
from collections import defaultdict
from typing import List, Dict
import os

# Configuration
SPOTIPY_CLIENT_ID = os.getenv('SPOTIPY_CLIENT_ID', 'b3edfd7466254e7780b2679b539d7382')
SPOTIPY_CLIENT_SECRET = os.getenv('SPOTIPY_CLIENT_SECRET', '56956b995dd3440faf4bbc39cdf1e218')
SPOTIPY_REDIRECT_URI = os.getenv('SPOTIPY_REDIRECT_URI', 'http://127.0.0.1:8888/callback')

# Rate limiting settings
BATCH_SIZE = 50  # Spotify API limit for most endpoints
RETRY_DELAY = 5  # seconds
MAX_RETRIES = 3
RATE_LIMIT_DELAY = 0.5  # seconds between requests


class SpotifyArtistCounter:
    def __init__(self):
        """Initialize Spotify client with OAuth"""
        scope = "user-library-read"

        self.sp = spotipy.Spotify(auth_manager=SpotifyOAuth(
            client_id=SPOTIPY_CLIENT_ID,
            client_secret=SPOTIPY_CLIENT_SECRET,
            redirect_uri=SPOTIPY_REDIRECT_URI,
            scope=scope
        ))

        self.user_id = self.sp.current_user()['id']
        print(f"Autenticado como: {self.user_id}")

    def fetch_all_liked_songs(self) -> List[Dict]:
        """
        Fetch all liked songs with progress tracking and error recovery
        Handles large libraries (16500+) with proper pagination
        """
        liked_songs = []
        offset = 0
        progress_file = 'liked_songs_progress.json'

        # Try to resume from saved progress
        if os.path.exists(progress_file):
            print("Arquivo de progresso encontrado, retomando do último salvamento...")
            with open(progress_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                liked_songs = data['songs']
                offset = data['offset']
                print(f"Retomando do offset {offset}, já tenho {len(liked_songs)} músicas")

        print(f"Buscando músicas curtidas a partir do offset {offset}...")

        while True:
            retry_count = 0
            while retry_count < MAX_RETRIES:
                try:
                    results = self.sp.current_user_saved_tracks(limit=BATCH_SIZE, offset=offset)
                    break
                except Exception as e:
                    retry_count += 1
                    print(f"Erro ao buscar músicas (tentativa {retry_count}/{MAX_RETRIES}): {e}")
                    if retry_count >= MAX_RETRIES:
                        print("Máximo de tentativas atingido. Salvando progresso...")
                        self._save_progress(progress_file, liked_songs, offset)
                        raise
                    time.sleep(RETRY_DELAY * retry_count)

            if not results['items']:
                break

            liked_songs.extend(results['items'])
            offset += len(results['items'])

            print(f"Buscadas {offset} músicas até agora...")

            # Save progress every 500 songs
            if offset % 500 == 0:
                self._save_progress(progress_file, liked_songs, offset)
                print(f"Progresso salvo em {offset} músicas")

            if not results['next']:
                break

            time.sleep(RATE_LIMIT_DELAY)  # Rate limiting

        print(f"\nTotal de músicas curtidas buscadas: {len(liked_songs)}")

        # Save final progress
        self._save_progress(progress_file, liked_songs, offset)

        return liked_songs

    def _save_progress(self, filename: str, songs: List[Dict], offset: int):
        """Save progress to file for recovery"""
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump({'songs': songs, 'offset': offset}, f, ensure_ascii=False)

    def count_songs_by_artist(self, liked_songs: List[Dict]) -> Dict[str, Dict]:
        """
        Count how many songs you have from each artist
        Returns dict with artist info and song count
        """
        artist_counts = defaultdict(lambda: {
            'count': 0,
            'songs': [],
            'artist_id': None,
            'artist_name': None
        })

        print("\nContabilizando músicas por artista...")
        print(f"Processando {len(liked_songs)} músicas...")

        for idx, item in enumerate(liked_songs):
            if idx % 500 == 0 and idx > 0:
                print(f"Processando música {idx}/{len(liked_songs)}")

            track = item['track']

            # Process all artists for this track
            for artist in track['artists']:
                artist_id = artist['id']
                artist_name = artist['name']

                # Initialize or update artist data
                if artist_counts[artist_id]['artist_id'] is None:
                    artist_counts[artist_id]['artist_id'] = artist_id
                    artist_counts[artist_id]['artist_name'] = artist_name

                # Increment count and add song info
                artist_counts[artist_id]['count'] += 1
                artist_counts[artist_id]['songs'].append({
                    'name': track['name'],
                    'id': track['id'],
                    'uri': track['uri']
                })

        print(f"\nContabilização completa!")
        print(f"Total de artistas únicos: {len(artist_counts)}")

        return dict(artist_counts)

    def display_and_save_results(self, artist_counts: Dict[str, Dict]):
        """
        Display results sorted by song count and save to file
        """
        # Sort artists by song count (descending)
        sorted_artists = sorted(
            artist_counts.items(),
            key=lambda x: x[1]['count'],
            reverse=True
        )

        print("\n" + "=" * 80)
        print("CONTAGEM DE MÚSICAS POR ARTISTA")
        print("=" * 80)
        print(f"\n{'#':<5} {'Artista':<40} {'Músicas':>10}")
        print("-" * 80)

        # Display top artists
        for idx, (artist_id, data) in enumerate(sorted_artists[:50], 1):
            artist_name = data['artist_name']
            count = data['count']

            # Truncate long names
            if len(artist_name) > 37:
                artist_name = artist_name[:34] + "..."

            print(f"{idx:<5} {artist_name:<40} {count:>10}")

        if len(sorted_artists) > 50:
            print(f"\n... e mais {len(sorted_artists) - 50} artistas")

        # Calculate statistics
        total_songs = sum(data['count'] for _, data in sorted_artists)
        avg_songs = total_songs / len(sorted_artists) if sorted_artists else 0

        print("\n" + "=" * 80)
        print("ESTATÍSTICAS")
        print("=" * 80)
        print(f"Total de músicas: {total_songs}")
        print(f"Total de artistas únicos: {len(sorted_artists)}")
        print(f"Média de músicas por artista: {avg_songs:.2f}")
        print(f"Artista com mais músicas: {sorted_artists[0][1]['artist_name']} ({sorted_artists[0][1]['count']} músicas)")

        # Save detailed results to JSON
        results = {
            'statistics': {
                'total_songs': total_songs,
                'total_artists': len(sorted_artists),
                'average_songs_per_artist': round(avg_songs, 2),
                'top_artist': {
                    'name': sorted_artists[0][1]['artist_name'],
                    'count': sorted_artists[0][1]['count']
                }
            },
            'artists': [
                {
                    'rank': idx,
                    'artist_id': artist_id,
                    'artist_name': data['artist_name'],
                    'song_count': data['count'],
                    'songs': data['songs']
                }
                for idx, (artist_id, data) in enumerate(sorted_artists, 1)
            ]
        }

        # Save to JSON file
        output_file = 'artist_song_counts.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        print(f"\nResultados detalhados salvos em: {output_file}")

        # Save a simplified CSV for easy viewing
        csv_file = 'artist_song_counts.csv'
        with open(csv_file, 'w', encoding='utf-8') as f:
            f.write("Rank,Artista,Quantidade de Músicas\n")
            for idx, (artist_id, data) in enumerate(sorted_artists, 1):
                # Escape commas in artist names
                artist_name = data['artist_name'].replace(',', ';')
                f.write(f"{idx},{artist_name},{data['count']}\n")

        print(f"Resumo em CSV salvo em: {csv_file}")

        return results

    def run(self):
        """Main execution function"""
        print("=" * 80)
        print("CONTADOR DE MÚSICAS POR ARTISTA - SPOTIFY")
        print("=" * 80)

        # Step 1: Fetch all liked songs
        liked_songs = self.fetch_all_liked_songs()

        if not liked_songs:
            print("Nenhuma música curtida encontrada!")
            return

        # Step 2: Count songs by artist
        artist_counts = self.count_songs_by_artist(liked_songs)

        # Step 3: Display and save results
        results = self.display_and_save_results(artist_counts)

        print("\n" + "=" * 80)
        print("CONCLUÍDO!")
        print("=" * 80)
        print("\nArquivos gerados:")
        print("  - artist_song_counts.json (dados completos com lista de músicas)")
        print("  - artist_song_counts.csv (resumo para Excel/Sheets)")
        print("  - liked_songs_progress.json (cache para retomar se necessário)")


def main():
    """Main entry point"""
    print("\n🎵 Contador de Músicas por Artista - Spotify 🎵\n")
    print("INSTRUÇÕES DE CONFIGURAÇÃO:")
    print("1. Vá para https://developer.spotify.com/dashboard")
    print("2. Crie um app e obtenha seu Client ID e Client Secret")
    print("3. Configure o redirect URI para: http://localhost:8888/callback")
    print("4. Configure as variáveis de ambiente ou atualize o código:")
    print("   export SPOTIPY_CLIENT_ID='seu_client_id'")
    print("   export SPOTIPY_CLIENT_SECRET='seu_client_secret'")
    print("\n" + "=" * 80)

    if SPOTIPY_CLIENT_ID == 'YOUR_CLIENT_ID':
        print("\n⚠️  AVISO: Por favor, atualize suas credenciais da API do Spotify!")
        print("Edite o script ou configure as variáveis de ambiente antes de executar.\n")
        return

    try:
        counter = SpotifyArtistCounter()
        counter.run()

    except Exception as e:
        print(f"\n❌ Erro: {e}")
        print("\nO progresso foi salvo. Você pode executar o script novamente para retomar.")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()