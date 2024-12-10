# Python Challenge

You have a shapefile of lakes in the Omineca Region called Management Objectives. It was created by using the Omineca Wildlife Habitat Areas clipping the freshwater atlas layer.
This shapefile has 7 tables related to it. The tables are all related to the Management Objectives with Waterbody ID as the primary key. 
Unfortunately the tables contain data not just for Omineca, but the whole Province of BC. The client wants you to remove all data that is not in the Omineca region so that the 
related tables only contain the data for Omineca. If all the data was spatialized, you could use the Omineca Polygon to clip the related table data. However, the data
isn't spatialized so you will need to use python to clean the data. 