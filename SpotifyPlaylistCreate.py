"""
Spotify Playlist Organizer
Handles large libraries (16500+ songs) with rate limiting and error recovery
Creates genre-based playlists from your liked songs
"""

import spotipy
from spotipy.oauth2 import SpotifyOAuth
import time
import json
from collections import defaultdict
from typing import List, Dict, Set
import os

# Configuration
SPOTIPY_CLIENT_ID = os.getenv('SPOTIPY_CLIENT_ID', 'YOUR_CLIENT_ID')
SPOTIPY_CLIENT_SECRET = os.getenv('SPOTIPY_CLIENT_SECRET', 'YOUR_CLIENT_SECRET')
SPOTIPY_REDIRECT_URI = os.getenv('SPOTIPY_REDIRECT_URI', 'http://localhost:8888/callback')

# Rate limiting settings
BATCH_SIZE = 50  # Spotify API limit for most endpoints
RETRY_DELAY = 5  # seconds
MAX_RETRIES = 3
RATE_LIMIT_DELAY = 0.5  # seconds between requests

# Genre mapping - main genres with their subgenres
GENRE_GROUPS = {
    'Rock': ['rock', 'hard rock', 'classic rock', 'alternative rock', 'alt rock', 
             'indie rock', 'progressive rock', 'punk rock', 'garage rock', 'psychedelic rock',
             'folk rock', 'rock-and-roll', 'southern rock', 'art rock', 'glam rock',
             'post-rock', 'math rock', 'noise rock', 'grunge', 'britpop', 'rock pop',
             'pop rock', 'soft rock'],
    
    'Metal': ['metal', 'heavy metal', 'black metal', 'death metal', 'thrash metal',
              'doom metal', 'power metal', 'progressive metal', 'symphonic metal',
              'folk metal', 'viking metal', 'melodic death metal', 'metalcore',
              'deathcore', 'groove metal', 'nu metal', 'industrial metal', 'sludge metal',
              'stoner metal', 'gothic metal', 'speed metal'],
    
    'Electronic': ['electronic', 'edm', 'house', 'techno', 'trance', 'dubstep',
                   'drum and bass', 'dnb', 'ambient', 'downtempo', 'electro',
                   'synthwave', 'vaporwave', 'future bass', 'chillwave', 'idm',
                   'breakbeat', 'garage', 'uk garage', 'bass', 'trap', 'future garage',
                   'electronica', 'synth-pop', 'synthpop', 'electropop'],
    
    'Hip Hop': ['hip hop', 'rap', 'trap', 'conscious hip hop', 'gangsta rap',
                'underground hip hop', 'east coast hip hop', 'west coast hip hop',
                'southern hip hop', 'boom bap', 'lo-fi hip hop', 'cloud rap',
                'drill', 'grime', 'uk hip hop', 'experimental hip hop'],
    
    'Pop': ['pop', 'electropop', 'dance pop', 'synth-pop', 'indie pop', 'art pop',
            'dream pop', 'chamber pop', 'baroque pop', 'k-pop', 'j-pop', 'bubblegum pop',
            'teen pop', 'latin pop', 'europop'],
    
    'Jazz': ['jazz', 'bebop', 'cool jazz', 'free jazz', 'fusion', 'jazz fusion',
             'smooth jazz', 'contemporary jazz', 'latin jazz', 'soul jazz',
             'hard bop', 'modal jazz', 'swing', 'big band', 'gypsy jazz', 'nu jazz'],
    
    'Classical': ['classical', 'baroque', 'romantic', 'contemporary classical',
                  'minimalism', 'opera', 'orchestral', 'chamber music', 'symphonic',
                  'piano', 'choral', 'modern classical', 'neoclassical'],
    
    'R&B/Soul': ['r&b', 'soul', 'neo soul', 'contemporary r&b', 'funk', 'motown',
                 'disco', 'quiet storm', 'new jack swing', 'alternative r&b'],
    
    'Country': ['country', 'country pop', 'country rock', 'alternative country',
                'bluegrass', 'outlaw country', 'honky tonk', 'americana',
                'folk country', 'contemporary country'],
    
    'Indie/Alternative': ['indie', 'alternative', 'indie folk', 'indie pop', 'indie rock',
                          'lo-fi', 'bedroom pop', 'shoegaze', 'noise pop', 'indie electronic',
                          'folktronica', 'indie dance', 'new wave', 'post-punk'],
    
    'Reggae/Dancehall': ['reggae', 'dancehall', 'dub', 'ska', 'rocksteady', 'reggaeton',
                         'roots reggae', 'lovers rock', 'ragga'],
    
    'Latin': ['latin', 'salsa', 'bachata', 'merengue', 'cumbia', 'reggaeton',
              'banda', 'mariachi', 'ranchera', 'tango', 'bossa nova', 'samba',
              'flamenco', 'latin jazz', 'latin pop', 'tejano', 'urbano latino'],
    
    'Folk/Acoustic': ['folk', 'acoustic', 'singer-songwriter', 'traditional folk',
                      'contemporary folk', 'freak folk', 'folk pop', 'celtic',
                      'bluegrass', 'americana'],
    
    'Blues': ['blues', 'delta blues', 'chicago blues', 'electric blues',
              'blues rock', 'country blues', 'soul blues', 'british blues'],
    
    'World': ['world', 'afrobeat', 'african', 'indian', 'middle eastern',
              'asian', 'celtic', 'klezmer', 'bollywood', 'flamenco',
              'balkan', 'traditional', 'ethnic'],
    
    'Punk': ['punk', 'hardcore punk', 'pop punk', 'ska punk', 'post-punk',
             'anarcho-punk', 'crust punk', 'street punk', 'emo', 'screamo',
             'powerviolence', 'hardcore'],
    
    'Soundtrack/Score': ['soundtrack', 'score', 'film score', 'game soundtrack',
                         'musical', 'theatre', 'video game music'],
}


class SpotifyPlaylistOrganizer:
    def __init__(self):
        """Initialize Spotify client with OAuth"""
        scope = "user-library-read playlist-modify-public playlist-modify-private"
        
        self.sp = spotipy.Spotify(auth_manager=SpotifyOAuth(
            client_id=SPOTIPY_CLIENT_ID,
            client_secret=SPOTIPY_CLIENT_SECRET,
            redirect_uri=SPOTIPY_REDIRECT_URI,
            scope=scope
        ))
        
        self.user_id = self.sp.current_user()['id']
        print(f"Authenticated as: {self.user_id}")
        
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
            print("Found progress file, resuming from last save...")
            with open(progress_file, 'r') as f:
                data = json.load(f)
                liked_songs = data['songs']
                offset = data['offset']
                print(f"Resuming from offset {offset}, already have {len(liked_songs)} songs")
        
        print(f"Fetching liked songs starting from offset {offset}...")
        
        while True:
            retry_count = 0
            while retry_count < MAX_RETRIES:
                try:
                    results = self.sp.current_user_saved_tracks(limit=BATCH_SIZE, offset=offset)
                    break
                except Exception as e:
                    retry_count += 1
                    print(f"Error fetching songs (attempt {retry_count}/{MAX_RETRIES}): {e}")
                    if retry_count >= MAX_RETRIES:
                        print("Max retries reached. Saving progress...")
                        self._save_progress(progress_file, liked_songs, offset)
                        raise
                    time.sleep(RETRY_DELAY * retry_count)
            
            if not results['items']:
                break
                
            liked_songs.extend(results['items'])
            offset += len(results['items'])
            
            print(f"Fetched {offset} songs so far...")
            
            # Save progress every 500 songs
            if offset % 500 == 0:
                self._save_progress(progress_file, liked_songs, offset)
                print(f"Progress saved at {offset} songs")
            
            if not results['next']:
                break
                
            time.sleep(RATE_LIMIT_DELAY)  # Rate limiting
        
        print(f"\nTotal liked songs fetched: {len(liked_songs)}")
        
        # Save final progress
        self._save_progress(progress_file, liked_songs, offset)
        
        return liked_songs
    
    def _save_progress(self, filename: str, songs: List[Dict], offset: int):
        """Save progress to file for recovery"""
        with open(filename, 'w') as f:
            json.dump({'songs': songs, 'offset': offset}, f)
    
    def get_track_genres(self, track: Dict) -> List[str]:
        """
        Get genres for a track by checking artist genres
        With retry logic and rate limiting
        """
        genres = []
        
        try:
            for artist in track['track']['artists']:
                retry_count = 0
                while retry_count < MAX_RETRIES:
                    try:
                        artist_info = self.sp.artist(artist['id'])
                        genres.extend(artist_info['genres'])
                        time.sleep(RATE_LIMIT_DELAY)
                        break
                    except Exception as e:
                        retry_count += 1
                        if retry_count >= MAX_RETRIES:
                            print(f"Failed to get genres for artist {artist['name']}: {e}")
                            break
                        time.sleep(RETRY_DELAY)
        except Exception as e:
            print(f"Error getting track genres: {e}")
        
        return genres
    
    def categorize_by_genre(self, liked_songs: List[Dict]) -> Dict[str, List[str]]:
        """
        Categorize songs by main genre groups
        Returns dict mapping genre group to list of track URIs
        """
        genre_playlists = defaultdict(list)
        unmatched_tracks = []
        
        print("\nCategorizing songs by genre...")
        print("This will take a while for 16500+ songs...")
        
        for idx, item in enumerate(liked_songs):
            if idx % 100 == 0:
                print(f"Processing song {idx + 1}/{len(liked_songs)}")
            
            track = item['track']
            track_uri = track['uri']
            
            # Get all genres for this track
            track_genres = self.get_track_genres(item)
            track_genres_lower = [g.lower() for g in track_genres]
            
            # Find which main genre group(s) this track belongs to
            matched = False
            for main_genre, subgenres in GENRE_GROUPS.items():
                if any(subgenre in genre for genre in track_genres_lower for subgenre in subgenres):
                    genre_playlists[main_genre].append(track_uri)
                    matched = True
            
            if not matched:
                unmatched_tracks.append({
                    'uri': track_uri,
                    'name': track['name'],
                    'artist': track['artists'][0]['name'],
                    'genres': track_genres
                })
        
        # Save categorization results
        with open('genre_categorization.json', 'w') as f:
            json.dump({
                'genre_playlists': {k: v for k, v in genre_playlists.items()},
                'unmatched': unmatched_tracks
            }, f, indent=2)
        
        print(f"\nCategorization complete!")
        for genre, tracks in sorted(genre_playlists.items(), key=lambda x: len(x[1]), reverse=True):
            print(f"  {genre}: {len(tracks)} songs")
        print(f"  Unmatched: {len(unmatched_tracks)} songs")
        
        return genre_playlists
    
    def create_playlist(self, name: str, description: str, track_uris: List[str]):
        """
        Create a playlist and add tracks in batches
        Handles Spotify's 100-track limit per request AND 10,000 song playlist limit
        Creates multiple playlists (Part 1, Part 2, etc.) if needed
        """
        MAX_PLAYLIST_SIZE = 10000
        total_tracks = len(track_uris)
        
        # Calculate how many playlists we need
        num_playlists = (total_tracks + MAX_PLAYLIST_SIZE - 1) // MAX_PLAYLIST_SIZE
        
        created_playlists = []
        
        for playlist_num in range(num_playlists):
            # Calculate which tracks go in this playlist
            start_idx = playlist_num * MAX_PLAYLIST_SIZE
            end_idx = min(start_idx + MAX_PLAYLIST_SIZE, total_tracks)
            playlist_tracks = track_uris[start_idx:end_idx]
            
            # Determine playlist name
            if num_playlists > 1:
                playlist_name = f"{name} (Part {playlist_num + 1})"
                playlist_desc = f"{description} - Part {playlist_num + 1} of {num_playlists}"
            else:
                playlist_name = name
                playlist_desc = description
            
            print(f"\nCreating playlist: {playlist_name}")
            print(f"  Will contain {len(playlist_tracks)} songs")
            
            # Create playlist
            retry_count = 0
            while retry_count < MAX_RETRIES:
                try:
                    playlist = self.sp.user_playlist_create(
                        self.user_id,
                        playlist_name,
                        public=False,
                        description=playlist_desc
                    )
                    break
                except Exception as e:
                    retry_count += 1
                    print(f"Error creating playlist (attempt {retry_count}/{MAX_RETRIES}): {e}")
                    if retry_count >= MAX_RETRIES:
                        raise
                    time.sleep(RETRY_DELAY * retry_count)
            
            playlist_id = playlist['id']
            
            # Add tracks in batches of 100 (Spotify limit)
            for i in range(0, len(playlist_tracks), 100):
                batch = playlist_tracks[i:i + 100]
                retry_count = 0
                
                while retry_count < MAX_RETRIES:
                    try:
                        self.sp.playlist_add_items(playlist_id, batch)
                        print(f"  Added {min(i + 100, len(playlist_tracks))}/{len(playlist_tracks)} tracks")
                        break
                    except Exception as e:
                        retry_count += 1
                        print(f"Error adding tracks (attempt {retry_count}/{MAX_RETRIES}): {e}")
                        if retry_count >= MAX_RETRIES:
                            print(f"Failed to add batch starting at {i}")
                            break
                        time.sleep(RETRY_DELAY * retry_count)
                
                time.sleep(RATE_LIMIT_DELAY)
            
            print(f"✓ Playlist '{playlist_name}' created with {len(playlist_tracks)} songs")
            print(f"  Playlist URL: https://open.spotify.com/playlist/{playlist_id}")
            
            created_playlists.append(playlist_id)
        
        return created_playlists
    
    def create_all_playlists(self):
        """Main orchestration function"""
        print("=" * 60)
        print("SPOTIFY PLAYLIST ORGANIZER")
        print("=" * 60)
        
        # Step 1: Fetch all liked songs
        liked_songs = self.fetch_all_liked_songs()
        
        if not liked_songs:
            print("No liked songs found!")
            return
        
        # Step 2: Create "All Liked Songs" playlist(s)
        # Note: If you have more than 10,000 songs, multiple playlists will be created
        all_track_uris = [item['track']['uri'] for item in liked_songs]
        
        if len(all_track_uris) > 10000:
            print(f"\n⚠️  You have {len(all_track_uris)} liked songs!")
            print(f"Spotify's limit is 10,000 songs per playlist.")
            print(f"Will create {(len(all_track_uris) + 9999) // 10000} playlists for all your liked songs.\n")
        
        self.create_playlist(
            name="All Liked Songs",
            description=f"All {len(all_track_uris)} of my liked songs",
            track_uris=all_track_uris
        )
        
        # Step 3: Categorize by genre
        genre_playlists = self.categorize_by_genre(liked_songs)
        
        # Step 4: Create genre-based playlists
        print("\n" + "=" * 60)
        print("CREATING GENRE PLAYLISTS")
        print("=" * 60)
        
        for genre, track_uris in sorted(genre_playlists.items(), key=lambda x: len(x[1]), reverse=True):
            if len(track_uris) > 0:  # Only create if there are songs
                if len(track_uris) > 10000:
                    print(f"\n⚠️  {genre} has {len(track_uris)} songs (over 10k limit)")
                    print(f"Will create multiple playlists for this genre.\n")
                
                self.create_playlist(
                    name=f"{genre} - My Collection",
                    description=f"All my {genre.lower()} songs from my liked tracks ({len(track_uris)} songs)",
                    track_uris=track_uris
                )
                time.sleep(1)  # Extra delay between playlist creations
        
        print("\n" + "=" * 60)
        print("ALL DONE!")
        print("=" * 60)
        print(f"\nSuccessfully organized your {len(all_track_uris)} liked songs!")
        print("Check 'genre_categorization.json' for detailed categorization results")
        print("including unmatched songs that you can manually categorize.")


def main():
    """Main entry point"""
    print("\n🎵 Spotify Playlist Organizer 🎵\n")
    print("SETUP INSTRUCTIONS:")
    print("1. Go to https://developer.spotify.com/dashboard")
    print("2. Create an app and get your Client ID and Client Secret")
    print("3. Set redirect URI to: http://localhost:8888/callback")
    print("4. Set environment variables or update the code:")
    print("   export SPOTIPY_CLIENT_ID='your_client_id'")
    print("   export SPOTIPY_CLIENT_SECRET='your_client_secret'")
    print("\n" + "=" * 60)
    
    if SPOTIPY_CLIENT_ID == 'YOUR_CLIENT_ID':
        print("\n⚠️  WARNING: Please update your Spotify API credentials!")
        print("Edit the script or set environment variables before running.\n")
        return
    
    try:
        organizer = SpotifyPlaylistOrganizer()
        organizer.create_all_playlists()
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        print("\nProgress has been saved. You can run the script again to resume.")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
