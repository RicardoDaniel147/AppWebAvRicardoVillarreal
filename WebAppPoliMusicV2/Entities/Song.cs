using System;
using System.Collections.Generic;
using System.Web;

namespace WebAppPoliMusicV2.Entities
{
    public class Song
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public int Plays { get; set; }
        public string Genre { get; set; }

        public Song(int id, string name, string path, int plays, string genre)
        {
            ID = id;
            Name = name;
            Path = path;
            Plays = plays;
            Genre = genre;
        }

        public Song()
        {
        }
    }
}
