package org.example;

public class Element {
    String name;
    int count;

    public Element(String name, int count){
        this.name = name;
        this.count = count;
    }

    public String getName(){
        return this.name;
    }

    public Integer getCount(){
        return this.count;
    }

    @Override
    public String toString() {
        return "\n"+name+" "+count;
    }
}

