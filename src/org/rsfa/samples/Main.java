package org.rsfa.samples;

public class Main {

    public static void main(String[] args) {
        if (args.length < 2) {
            System.err.println("ERROR: no input / ouput files specified.");
            return;
        }
        Extractor extractor = new Extractor();
        extractor.setInputFile(args[0]);
        extractor.setOutputFile(args[1]);
        extractor.read();
        extractor.extract();
        extractor.write();
    }
}
