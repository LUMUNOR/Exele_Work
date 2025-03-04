package org.example;

import java.util.ArrayList;

public class Result {
    Element sozv;
    ArrayList <Element> signal;
    int coeffient;

    public Result(Element sozv,   ArrayList <Element> signal,int coeffient){
        this.sozv = sozv;
        this.signal = signal;
        this.coeffient =coeffient;
    }

    public Element getSozv(){
        return this.sozv;
    }

    public ArrayList<Element> getSignal(){
        return this.signal;
    }

    public int getCoffient(){
        return this.coeffient;
    }

   }

