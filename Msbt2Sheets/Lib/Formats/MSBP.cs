using System.Drawing;
using System.Text;
using Msbt2Sheets.Lib.Formats.FileComponents;
using Msbt2Sheets.Lib.Utils;

namespace Msbt2Sheets.Lib.Formats;

public class MSBP : GeneralFile
{
    public Dictionary<string, Color> Colors = new();
    public List<AttributeInfo> AttributeInfos = new();
    public List<TagGroup> TagGroups = new();
    public List<Style> Styles = new();
    public List<string> SourceFileNames = new();
    
    public bool HasCLR1 = false;
    public bool HasCLB1 = false;
    public bool HasATI2 = false;
    public bool HasALB1 = false;
    public bool HasALI2 = false;
    public bool HasTGG2 = false;
    public bool HasTAG2 = false;
    public bool HasTGP2 = false;
    public bool HasTGL2 = false;
    public bool HasSYL3 = false;
    public bool HasSLB1 = false;
    public bool HasCTI1 = false;
    
    public Header Header = new();

    public MSBP()
    {
    }

    public MSBP(Stream fileStream)
    {
        MemoryStream ms = new MemoryStream();
        fileStream.CopyTo(ms);
        FileReader reader = new(ms);

        CLR1 clr1 = new();
        CLB1 clb1 = new();
        ATI2 ati2 = new();
        ALB1 alb1 = new();
        ALI2 ali2 = new();
        TGG2 tgg2 = new();
        TAG2 tag2 = new();
        TGP2 tgp2 = new();
        TGL2 tgl2 = new();
        SYL3 syl3 = new();
        SLB1 slb1 = new();
        CTI1 cti1 = new();
        
        Header = new(reader);

        for (int i = 0; i < Header.SectionCount; i++)
        {
            string sectionMagic = reader.ReadString(4, Encoding.ASCII);
            uint sectionSize = reader.ReadUInt32();
            reader.Skip(8);
            long startPosition = reader.Position;

            switch (sectionMagic)
            {
                case "CLR1":
                    HasCLR1 = true;
                    clr1 = new(reader);
                    break;
                case "CLB1":
                    HasCLB1 = true;
                    clb1 = new(reader);
                    break;
                case "ATI2":
                    HasATI2 = true;
                    ati2 = new(reader);
                    break;
                case "ALB1":
                    HasALB1 = true;
                    alb1 = new(reader);
                    break;
                case "ALI2":
                    HasALI2 = true;
                    ali2 = new(reader);
                    break;
                case "TGG2":
                    HasTGG2 = true;
                    tgg2 = new(reader);
                    break;
                case "TAG2":
                    HasTAG2 = true;
                    tag2 = new(reader);
                    break;
                case "TGP2":
                    HasTGP2 = true;
                    tgp2 = new(reader);
                    break;
                case "TGL2":
                    HasTGL2 = true;
                    tgl2 = new(reader);
                    break;
                case "SYL3":
                    HasSYL3 = true;
                    syl3 = new(reader);
                    break;
                case "SLB1":
                    HasSLB1 = true;
                    slb1 = new(reader);
                    break;
                case "CTI1":
                    HasCTI1 = true;
                    cti1 = new(reader);
                    break;
                default:
                    throw new InvalidDataException($"Unknown section magic!");
            }
            
            reader.JumpTo(startPosition);
            reader.Skip((int)sectionSize);
            reader.Align(0x10);
        }

        if (HasCLR1)
        {
            for (int i = 0; i < clr1.Colors.Count; i++)
            {
                if (i < clb1.ColorLabels.Count)
                {
                    Colors.Add(clb1.ColorLabels[i], clr1.Colors[i]);
                }
                else
                {
                    Colors.Add(i.ToString(), clr1.Colors[i]);
                }
            }
        }

        if (HasATI2)
        {
            for (int i = 0; i < alb1.AttributeLabels.Count; i++)
            {
                AttributeInfos.Add(new AttributeInfo(alb1.AttributeLabels[i], ati2.AttributeTypes[i], ati2.AttributeOffsets[i], ali2.AttributeLists[ati2.AttributeListIds[i]]));
            }
        }

        if (HasTGG2)
        {
            for (int i = 0; i < tgg2.TagGroupNames.Count; i++)
            {
                TagGroup group = new();
                group.Name = tgg2.TagGroupNames[i];
                group.Id = tgg2.TagGroupIds[i];
                foreach (var tagId in tgg2.TagIdLists[i])
                {
                    TagType tag = new();
                    tag.Name = tag2.TagNames[tagId];
                    foreach (var parameterId in tag2.TagArgumentIdLists[tagId])
                    {
                        TagParameter param = new();
                        param.Name = tgp2.ParameterNames[parameterId];
                        param.Type = (ParamType)tgp2.ParameterTypes[parameterId];
                        param.ListItemIds = tgp2.ListItemIdLists[parameterId];
                        foreach (var listItemId in param.ListItemIds)
                        {
                            param.List.Add(tgl2.ListItems[listItemId]);
                        }
                        
                        tag.Parameters.Add(param);
                    }
                    
                    group.Tags.Add(tag);
                }
                
                TagGroups.Add(group);
            }
        }

        if (HasSYL3)
        {
            Styles = syl3.Styles;
            for (int i = 0; i < Styles.Count; i++)
            {
                Styles[i].Name = slb1.StyleLabels[i];
            }
        }

        if (HasCTI1)
        {
            SourceFileNames = cti1.FileNames;
        }

        if (TagGroups.Count == 0)
        {
            TagGroups = BaseMSBP.TagGroups;
        }
    }

    internal class CLR1
    {
        public List<Color> Colors = new();
        
        public CLR1() {}

        public CLR1(FileReader reader)
        {
            uint colorCount = reader.ReadUInt32();

            for (uint i = 0; i < colorCount; i++)
            {
                byte[] colorBytes = reader.ReadBytes(4);
                Colors.Add(Color.FromArgb(colorBytes[3], colorBytes[0], colorBytes[1], colorBytes[2]));
            }
        }
    }

    internal class CLB1
    {
        public List<string> ColorLabels = new();
        
        public CLB1() {}

        public CLB1(FileReader reader)
        {
            ColorLabels = ReadLabels(reader);
        }
    }
    
    internal class ATI2
    {
        public List<byte> AttributeTypes = new();
        public List<uint> AttributeOffsets = new();
        public List<ushort> AttributeListIds = new();
        
        public ATI2() {}

        public ATI2(FileReader reader)
        {
            uint attributeCount = reader.ReadUInt32();
            for (uint i = 0; i < attributeCount; i++)
            {
                byte type = reader.ReadByte();
                reader.Skip(1);
                ushort listId = reader.ReadUInt16();
                uint offset = reader.ReadUInt32();
                
                AttributeTypes.Add(type);
                AttributeOffsets.Add(offset);
                AttributeListIds.Add(listId);
            }
        }
    }
    
    internal class ALB1
    {
        public List<string> AttributeLabels = new();
        
        public ALB1() {}

        public ALB1(FileReader reader)
        {
            AttributeLabels = ReadLabels(reader);
        }
    }
    
    internal class ALI2
    {
        public List<List<string>> AttributeLists = new();
            
        public ALI2() {}

        public ALI2(FileReader reader)
        {
            long startPosition = reader.Position;
            
            uint listCount = reader.ReadUInt32();
            List<uint> listOffsets = new();
            for (uint i = 0; i < listCount; i++)
            {
                listOffsets.Add(reader.ReadUInt32());
            }

            for (uint i = 0; i < listCount; i++)
            {
                List<string> list = new();
                
                uint listItemCount = reader.ReadUInt32At(startPosition + listOffsets[(int)i]);
                List<uint> listItemOffsets = new();
                for (uint j = 0; j < listItemCount; j++)
                {
                    listItemOffsets.Add(reader.ReadUInt32());
                }

                for (uint j = 0; j < listItemCount; j++)
                {
                    list.Add(reader.ReadTerminatedStringAt(startPosition + listOffsets[(int) i] + listItemOffsets[(int) j]));
                }
                
                AttributeLists.Add(list);
            }
        }
    }
    
    internal class TGG2
    {
        public List<string> TagGroupNames = new();
        public List<ushort> TagGroupIds = new();
        public List<List<ushort>> TagIdLists = new();
        
        public TGG2() {}

        public TGG2(FileReader reader)
        {
            long startPosition = reader.Position;
            ushort tagGroupCount = reader.ReadUInt16();
            reader.Skip(2);

            List<uint> tagGroupOffsets = new();
            for (int i = 0; i < tagGroupCount; i++)
            {
                tagGroupOffsets.Add(reader.ReadUInt32());
            }

            bool hasGroupIds = false;
            
            for (int i = 0; i < tagGroupCount; i++)
            {
                reader.JumpTo(startPosition + tagGroupOffsets[i]);
                
                ushort firstNumber = reader.ReadUInt16();
                ushort tagCount;
                ushort groupId;
                if (i == 0)
                {
                    if (firstNumber == 0)
                    {
                        hasGroupIds = true;
                    }
                    else
                    {
                        hasGroupIds = false;
                    }
                }

                if (hasGroupIds)
                {
                    groupId = firstNumber;
                    tagCount = reader.ReadUInt16();
                }
                else
                {
                    groupId = (ushort)i;
                    tagCount = firstNumber;
                }

                TagGroupIds.Add(groupId);
                
                List<ushort> tagIds = new();
                for (int j = 0; j < tagCount; j++)
                {
                    tagIds.Add(reader.ReadUInt16());
                }

                TagIdLists.Add(tagIds);
                TagGroupNames.Add(reader.ReadTerminatedString());
                reader.Skip(1);
            }
        }
    }
    
    internal class TAG2
    {
        public List<List<ushort>> TagArgumentIdLists = new();
        public List<string> TagNames = new();
        
        public TAG2() {}

        public TAG2(FileReader reader)
        {
            long startPosition = reader.Position;
            ushort tagCount = reader.ReadUInt16();
            reader.Skip(2);

            List<uint> tagOffsets = new();
            for (int i = 0; i < tagCount; i++)
            {
                tagOffsets.Add(reader.ReadUInt32());
            }

            for (int i = 0; i < tagCount; i++)
            {
                reader.JumpTo(startPosition + tagOffsets[i]);

                ushort tagArgumentCount = reader.ReadUInt16();
                List<ushort> tagArgumentIds = new();
                for (int j = 0; j < tagArgumentCount; j++)
                {
                    tagArgumentIds.Add(reader.ReadUInt16());
                }

                TagArgumentIdLists.Add(tagArgumentIds);
                TagNames.Add(reader.ReadTerminatedString());
            }
        }
    }
    
    internal class TGP2
    {
        public List<List<ushort>> ListItemIdLists = new();
        public List<string> ParameterNames = new();
        public List<byte> ParameterTypes = new();
        
        public TGP2() {}

        public TGP2(FileReader reader)
        {
            long startPosition = reader.Position;
            ushort parameterCount = reader.ReadUInt16();
            reader.Skip(2);
            
            List<uint> parameterOffsets = new();
            for (int i = 0; i < parameterCount; i++)
            {
                parameterOffsets.Add(reader.ReadUInt32());
            }

            for (int i = 0; i < parameterCount; i++)
            {
                reader.JumpTo(startPosition + parameterOffsets[i]);

                byte type = reader.ReadByte();
                ParameterTypes.Add(type);

                if (type == 9)
                {
                    reader.Skip(1);
                    
                    ushort listItemCount = reader.ReadUInt16();
                    List<ushort> listItemIds = new();
                    for (int j = 0; j < listItemCount; j++)
                    {
                        listItemIds.Add(reader.ReadUInt16());
                    }
                    
                    ListItemIdLists.Add(listItemIds);
                }
                else
                {
                    ListItemIdLists.Add(new List<ushort>());
                }
                
                ParameterNames.Add(reader.ReadTerminatedString());
            }
        }
    }
    
    internal class TGL2
    {
        public List<string> ListItems = new();
        
        public TGL2() {}

        public TGL2(FileReader reader)
        {
            long startPosition = reader.Position;
            ushort listItemCount = reader.ReadUInt16();
            reader.Skip(2);
            
            List<uint> listItemOffsets = new();
            for (int i = 0; i < listItemCount; i++)
            {
                listItemOffsets.Add(reader.ReadUInt32());
            }

            for (int i = 0; i < listItemCount; i++)
            {
                reader.JumpTo(startPosition + listItemOffsets[i]);
                
                ListItems.Add(reader.ReadTerminatedString());
            }
        }
    }
    
    internal class SYL3
    {
        public List<Style> Styles = new();
        
        public SYL3() {}

        public SYL3(FileReader reader)
        {
            uint styleCount = reader.ReadUInt32();

            for (int i = 0; i < styleCount; i++)
            {
                Styles.Add(new Style(reader.ReadInt32(), reader.ReadInt32(), reader.ReadInt32(), reader.ReadInt32()));
            }
        }
    }
    
    internal class SLB1
    {
        public List<string> StyleLabels = new();
        
        public SLB1() {}

        public SLB1(FileReader reader)
        {
            StyleLabels = ReadLabels(reader);
        }
    }
    
    internal class CTI1
    {
        public List<string> FileNames = new();
        
        public CTI1() {}

        public CTI1(FileReader reader)
        {
            long startPosition = reader.Position;
            uint fileNameCount = reader.ReadUInt32();
            
            List<uint> fileNameOffsets = new();
            for (int i = 0; i < fileNameCount; i++)
            {
                fileNameOffsets.Add(reader.ReadUInt32());
            }

            for (int i = 0; i < fileNameCount; i++)
            {
                reader.JumpTo(startPosition + fileNameOffsets[i]);
                
                FileNames.Add(reader.ReadTerminatedString());
            }
        }
    }
    
    public static MSBP BaseMSBP = new MSBP()
    {
        TagGroups = new List<TagGroup>()
        {
            new TagGroup()
            {
                Name = "System",
                Id = 0,
                Tags = new List<TagType>()
                {
                    new TagType()
                    {
                        Name = "Ruby",
                        Parameters = new List<TagParameter>()
                        {
                            new TagParameter()
                            {
                                Name = "rt",
                                Type = ParamType.String
                            }
                        }
                    },
                    new TagType()
                    {
                        Name = "Font",
                        Parameters = new List<TagParameter>()
                        {
                            new TagParameter()
                            {
                                Name = "face",
                                Type = ParamType.String
                            }
                        }
                    },
                    new TagType()
                    {
                        Name = "Size",
                        Parameters = new List<TagParameter>()
                        {
                            new TagParameter()
                            {
                                Name = "percent",
                                Type = ParamType.UInt16
                            },
                            new TagParameter()
                            {
                                Name = "size",
                                Type = ParamType.String
                            }
                        }
                    },
                    new TagType()
                    {
                        Name = "Color",
                        Parameters = new List<TagParameter>()
                        {
                            new TagParameter()
                            {
                                Name = "r",
                                Type = ParamType.UInt8
                            },
                            new TagParameter()
                            {
                                Name = "g",
                                Type = ParamType.UInt8
                            },
                            new TagParameter()
                            {
                                Name = "b",
                                Type = ParamType.UInt8
                            },
                            new TagParameter()
                            {
                                Name = "a",
                                Type = ParamType.UInt8
                            },
                            new TagParameter()
                            {
                                Name = "name",
                                Type = ParamType.String
                            },
                        }
                    },
                    new TagType()
                    {
                        Name = "PageBreak",
                        Parameters = new List<TagParameter>()
                    },
                    new TagType()
                    {
                        Name = "Reference",
                        Parameters = new List<TagParameter>()
                        {
                            new TagParameter()
                            {
                                Name = "mstxt",
                                Type = ParamType.String
                            },
                            new TagParameter()
                            {
                                Name = "label",
                                Type = ParamType.String
                            },
                            new TagParameter()
                            {
                                Name = "lang",
                                Type = ParamType.String
                            },
                        }
                    }
                }
            }
        }
    };
}