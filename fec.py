from dataclasses import dataclass
import csv
import sys

"""
CompteNum (numérique) + lettre 
CompAuxNum +Lettre 
CompAuxNum + Lettre +CompAuxLib = unique 
Split JournalCode 
pice ref +couple débit credit + EcritureLet (lettre) = unique, ligne suivante, la lettre 

"""

@dataclass
class FEC:

    journal_code: str
    compte_num: int
    compte_aux_num: int
    compte_aux_lib: str
    compte_aux_num_new: str

class FECParser:

    def __init__(self, path: str):
        self.path = path
        self.out_path = path.replace(".txt", ".out.txt")
        self.encoding = "ISO-8859-1"
        self.lettres: dict[int, str] = {}
        self.verifs: dict[str, str] = {}
        self.row = 1
        self.nb_change = 0
        self.nb_verifs = 0
        self.renums: set[tuple[int, str, str]] = set()

    def parse_row(self, row: dict[str, str], out):
        if row["CompAuxNum"] != "":
            compte_aux_num = row["CompAuxNum"]
            compte_num = int(row["CompteNum"])
            if compte_num in self.lettres:
                lettre = self.lettres[compte_num]
                compte_aux_num = lettre + compte_aux_num
                print(f"row: {self.row} => {compte_aux_num}")
                self.nb_change += 1
                row["CompAuxNum"] = compte_aux_num
                if compte_aux_num in self.verifs.keys():
                    if row["CompAuxLib"] != self.verifs[compte_aux_num]:
                        print(f"Bad verif: {compte_aux_num}, {row['CompAuxLib']} vs {self.verifs[compte_aux_num]}")
                        self.nb_verifs += 1
                        sys.exit(1)
                    else:
                        print(f"Verif ok {row['CompAuxLib']} == {self.verifs[compte_aux_num]}")
                else:
                    self.verifs[compte_aux_num] = row["CompAuxLib"]
                    self.renums.add((compte_num, compte_aux_num, row["CompAuxLib"]))
        out.write("\t".join(row.values()) + "\n")

    def parse(self):
        with open(self.path, encoding=self.encoding) as f:
            with open(self.out_path, "w", encoding=self.encoding) as out:
                reader = csv.DictReader(f, delimiter="\t")
                header = True
                for row in reader:
                    self.row += 1
                    if header:
                        header = False
                        out.write("\t".join(row.keys()) + "\n")
                    self.parse_row(row, out)

    def parse_lettre(self):
        with open("FECLettre2023.csv", encoding=self.encoding) as f:
            reader = csv.reader(f)
            for row in reader:
                self.lettres[int(row[1])] = row[0]

    def save_renums(self):
        with open("FECUniqueChanges.txt", "w", encoding=self.encoding) as f:
            f.write("CompteNum\tCompAuxNum\tCompAuxLib\n")
            for row in self.renums:
                f.write(f"{row[0]}\t{row[1]}\t{row[2]}\n")




if __name__ == '__main__':
    parser = FECParser("FEC2023.txt")
    parser.parse_lettre()
    parser.parse()
    parser.save_renums()
    print(f"Nb changes: {parser.nb_change}")
    print(f"Nb unique changes: {len(parser.renums)}")
    print(f"Nb Verif: {parser.nb_verifs}")

